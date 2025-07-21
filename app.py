from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import analyzer

app = FastAPI()

# 允许跨域
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/analyze")
async def analyze_grades(file: UploadFile = File(...)):
    # 保存上传文件
    contents = await file.read()
    with open("temp.xlsx", "wb") as f:
        f.write(contents)

    # 调用分析函数
    result = analyzer.analyze_single_file("temp.xlsx")
    return result