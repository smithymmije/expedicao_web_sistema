# mongo_users.py
import os
from pymongo import MongoClient
from werkzeug.security import check_password_hash, generate_password_hash

client = MongoClient(os.getenv("MONGO_URI"))
db = client.get_database("expedicao")
users = db["users"]

def usuario_por_id(uid: str):
    return users.find_one({"_id": uid})

def autenticar(username: str, password: str):
    u = users.find_one({"_id": username})
    if u and check_password_hash(u["password_hash"], password):
        return u
    return None