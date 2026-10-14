import os
import tempfile
from fastapi import FastAPI, UploadFile, File
from comparar_planilhas import comparar_planilhas
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()


app.add_middleware(
    CORSMiddleware,
    allow_origins=["null"],
    allow_credentials=False,
    allow_methods=["POST"],
    allow_headers=["*"],
)



@app.get("/")
def inicio():
    return {"mensagem": "Comparador de Planilhas v3.0"}

@app.post("/comparar")
async def comparar(
    arquivo_a: UploadFile = File(...),
    arquivo_b: UploadFile = File(...)
):
    with tempfile.TemporaryDirectory() as pasta_temporaria:
        caminho_a = os.path.join(pasta_temporaria, "planilha_a.xlsx")

        with open(caminho_a, "wb") as arquivo:
           arquivo.write(await arquivo_a.read())

        caminho_b = os.path.join(pasta_temporaria, "planilha_b.xlsx")

        with open(caminho_b, "wb") as arquivo:
           arquivo.write(await arquivo_b.read())
        
        caminho_relatorio = os.path.join(
            pasta_temporaria,
           "relatorio_comparacao.xlsx"
        )

        resultado = comparar_planilhas(
            arquivo_a=caminho_a,
            arquivo_b=caminho_b,
            caminho_relatorio=caminho_relatorio
        )

        if resultado is not True:
            return {"erro": resultado}

        with open(caminho_relatorio, "rb") as arquivo:
            conteudo = arquivo.read()

        return Response(
            content=conteudo,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
            "Content-Disposition": 'attachment; filename="relatorio_comparacao.xlsx"'
            }
        ) 