import azure
import azure.functions as func
import logging

bp = func.Blueprint()


@bp.function_name("AdditionalHTTPFunction")
@bp.route(route="brandnewroute")
def test_function(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Python HTTP processed a function")
    return func.HttpResponse("this works", status_code=200)
