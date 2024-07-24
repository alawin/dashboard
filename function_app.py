import azure.functions as func
import datetime
import json
import logging
import csv
import codecs
from additional_function import bp

app = func.FunctionApp()
app.register_blueprint(bp)


@app.function_name("FirstHTTPFunction")
@app.route(route="myroute", auth_level=func.AuthLevel.ANONYMOUS)
def test_function(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("a request ongoing")
    return func.HttpResponse("it works", status_code=200)


@app.function_name("SecondHTTPFunction")
@app.route(route="newroute", auth_level=func.AuthLevel.ANONYMOUS)
def test_function(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Second request ongoing")
    name = req.params.get("name")
    if name:
        message = f"hello, {name}, this works out just fine"
    else:
        message = f"this is function without name"
    return func.HttpResponse(message, status_code=200)


# @app.function_name("firstblobfunction")
# @app.blob_trigger(
#     arg_name="myblob", path="newcontainer/People.csv", connection="AzureWebJobsStorage"
# )
# def test_function(myblob: func.InputStream):
#     logging.info(
#         f"Python function triggered, will be stored in people.csv {myblob.name}"
#     )


# @app.function_name("BlobRead")
# @app.blob_trigger(
#     arg_name="readfile",
#     path="newcontainer/People2.csv",
#     connection="AzureWebJobsStorage",
# )
# def main(readfile: func.InputStream):
#     reader = csv.reader(codecs.iterdecode(readfile, "utf-8"))
#     for line in reader:
#         print(line)
