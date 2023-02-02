import json
obj1 = open("test.json" , "r") 
json_obj = json.load(obj1)

print(json_obj["book1"]["title"])