import json
import jsonschema
from jsonschema import validate


def get_schema():
    """This function loads the given schema available"""
    with open('slingtv_schema.json', 'r') as file:
        schema = json.load(file)
    return schema


def validate_json(json_data):
    """REF: https://json-schema.org/ """
    # Describe what kind of json you expect.
    execute_api_schema = get_schema()

    try:
        validate(instance=json_data, schema=execute_api_schema)
    except jsonschema.exceptions.ValidationError as err:
        print(err)
        err = "Given JSON data is InValid"
        return False, err

    message = "Given JSON data is Valid"
    return True, message


# Convert json to python object.
jsonData = json.loads('{"id" : 10,"name": "DonOfDen","contact_number":1234567890}')

d = {'metrics':[
    {'date' : '1-1-21', 'name' : 'Uses Impression Rate (%)', 'last_week_value' : 65.89, 'prior_week_value' : 66.51, 'mtd' : 66.35, 'qtd' : 67.06, 'target_week' : 12},
    {'date' : '1-1-21', 'name' : 'Fill rate, Ad only (%)', 'last_week_value' : 1.89, 'prior_week_value' : 11.51, 'mtd' : 111.35, 'qtd' : 1111.06, 'target_week' : 23},
    {'date' : '1-1-21', 'name' : 'Fill rate, Ad+Promo (%)', 'last_week_value' : 2.89, 'prior_week_value' : 22.51, 'mtd' : 222.35, 'qtd' : 2222.06, 'target_week' : 34},
    {'date' : '1-1-21', 'name' : 'ompleted Impression Rate (%)', 'last_week_value' : 3.89, 'prior_week_value' : 33.51, 'mtd' : 333.35, 'qtd' : 3333.06, 'target_week' : 45}
]}
for metric in d['metrics']:
    print(metric['name'])
    jsonData = json.loads(json.dumps(metric))

    # validate it
    is_valid, msg = validate_json(jsonData)
    print(msg)