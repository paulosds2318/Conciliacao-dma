from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from app.services import processar_e_enviar_planilha

app = FastAPI(title="Integração DMA Coral")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/v1/importar-contratos")
async def importar_contratos(
    file: UploadFile = File(...),
    acao: str = Form(...)
):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Apenas arquivos Excel (.xlsx, .xls) são aceitos.")

    try:
        conteudo = await file.read()
        resultado = processar_e_enviar_planilha(conteudo, acao)
        return resultado
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
