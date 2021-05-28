from typing import List
import json

import pandas as pd
from fastapi import FastAPI, UploadFile, File, Request
from starlette.middleware.cors import CORSMiddleware  #引入 CORS中间件模块
from starlette.responses import FileResponse
from pydantic import BaseModel
from dao import *


app = FastAPI()
origins = ["*"]

global daoObject

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"])

class fbaplan(BaseModel):
    planname: str
    eanlist: str
    amount: str
    addr: str

class product(BaseModel):
    row_num: int
    ean: int
    model_id:str
    asin:str
    name:str
    rmb_price:float
    ship_fee:float
    item_amount:int
    inner_length:float
    inner_width:float
    inner_height:float
    inner_weight:float
    out_length:float
    out_width:float
    out_height:float
    out_weight:float
    unit_per_box:int

# @app.get("/")
# async def serve_home(request: Request):
#     return templates.TemplateResponse("index.html", {"request": request})
@app.get("/")
def read_root():
    return {"Hello": "World"}


@app.post("/product")
def set_product(form:product):
    print(form)
    write_product_xlsx(Product(form,""),form.row_num)
    return 1

@app.get("/product/{id}")
def get_product(id:int):
    # if id
    sz=len(Dao([],[]).productList)
    if id>=sz:
        return None
    product=Dao([],[]).productList[id-1]
    tmp=product.__dict__
    tmp.pop('row')
    return tmp



@app.get("/products")
def getlist():
    list=Dao([],[])
    res={}
    res['products']=[]
    z=1
    for i in list.productList:
        res['products'].append({'num':z,'ean':i.ean,'model_id':i.model_id,'name':i.name})
        z+=1
    return res
@app.get("/items/{item_id}")
def read_item(item_id: int, q: str = None):
    return {"item_id": item_id, "q": q}

@app.get("/file")
def file():
    return FileResponse('./shipment.txt', filename='shipment.txt')


@app.get("/label")
def label():
    return FileResponse('./label.pdf', filename='label.pdf')
@app.get("/ship")
def ship():
    return FileResponse('./shipment.xlsx', filename='shipment.xlsx')

@app.get("/echo/{text}")
async def getEchoApi(text:str):
    return {"echo":text}

@app.get("/echo")
async def getEchoApi():
    return {"echo":"text"}
@app.get("/workflow")
def workflow():
    return FileResponse('./workflow.xlsx', filename='workflow.xlsx')



@app.post("/fbaplan")
async def plan(fba:fbaplan):
    daoObject= Dao(fba.eanlist.strip('\n').split('\n'),fba.amount.strip('\n').split('\n'))
    write_doc_plan(daoObject.products,daoObject.boxs,daoObject.ean_to_asin,fba.addr,fba.planname)
    makePdf(daoObject.eanlist)
    makeWorkFlowPlan(daoObject)
    return FileResponse('./shipment.txt', filename='shipment.txt')
@app.post('/upload')
async def recv_file(file: UploadFile = File(...)):
    file_data = await file.read()
    with open(file.filename,"wb") as fp:
        fp.write(file_data)
    fp.close()
    rt_msg = {
        "name": file.filename,
        "type": file.content_type
    }
    newDao = getFbaEanAndAmount(file.filename)
    daoObject = Dao(newDao[0],newDao[1])
    write_doc_shippment(daoObject,file.filename)
    makePdf(newDao[0])
    return rt_msg



if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='127.0.0.1', port=1099)
