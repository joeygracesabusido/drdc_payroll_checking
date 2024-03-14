from pymongo import MongoClient
import pymongo

def create_mongo_client():
    var_url = f"mongodb+srv://joeysabusido:genesis11@cluster0.r76lv.mongodb.net/drdc_payroll?retryWrites=true&w=majority"
    client = MongoClient(var_url, maxPoolSize=None)
    conn = client['drdc_payroll']

    # mongodb+srv://joeysabusido:<password>@cluster0.r76lv.mongodb.net/

    return conn 