from fastapi import FastAPI

app = FastAPI()

@app.get("/")
async def root():
    return {"message": "PPT MVP Backend"}

@app.post("/upload")
async def upload_template():
    return {"message": "Template upload endpoint - TBD"}

@app.post("/generate")
async def generate_slide():
    return {"message": "Slide generation endpoint - TBD"}