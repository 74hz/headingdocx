from fastapi import FastAPI

from app.api import router

app = FastAPI(title="headingdocx REST API")

app.include_router(router)
