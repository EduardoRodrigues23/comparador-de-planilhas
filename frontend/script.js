
const inputArquivoA = document.getElementById("arquivo_a");
const inputArquivoB = document.getElementById("arquivo_b");

const nomeArquivoA = document.getElementById("nome_arquivo_a");
const nomeArquivoB = document.getElementById("nome_arquivo_b");

inputArquivoA.addEventListener("change", function() {
    nomeArquivoA.textContent = inputArquivoA.files[0]?.name || "Clique para selecionar";
});

inputArquivoB.addEventListener("change", function() {
    nomeArquivoB.textContent = inputArquivoB.files[0]?.name || "Clique para selecionar";
});

const botaoComparar = document.getElementById("btn_comparar");

botaoComparar.addEventListener("click", function() {
   
const arquivoA = document.getElementById("arquivo_a").files[0];
const arquivoB = document.getElementById("arquivo_b").files[0];

if (!arquivoA || !arquivoB) {
    alert("Selecione as duas planilhas antes de comparar!");
    return;
}


const dados = new FormData();

dados.append("arquivo_a", arquivoA);
dados.append("arquivo_b", arquivoB);


botaoComparar.disabled = true;
botaoComparar.textContent = "Comparando...";
fetch("http://127.0.0.1:8000/comparar", {
    method: "POST",
    body: dados
})

.then(resposta => {
    if (!resposta.ok) {
        throw new Error("Erro ao comparar as planilhas.");
    }

    return resposta.blob();
})
.then(arquivo => {
    const url = URL.createObjectURL(arquivo);

    const link = document.createElement("a");
    link.href = url;
    link.download = "relatorio_comparacao.xlsx";

    document.body.appendChild(link);
    link.click();
    link.remove();

    URL.revokeObjectURL(url);
})


.catch(erro => {
    console.error("Erro na requisição:", erro);
    alert("Não foi possível comparar as planilhas.");
})
.finally(() => {
    botaoComparar.disabled = false;
    botaoComparar.textContent = "Comparar Planilhas";
});




});
