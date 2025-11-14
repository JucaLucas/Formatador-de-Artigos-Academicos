const fileInput = document.getElementById('fileInput');
const fileName = document.getElementById('fileName');
const fileNameResult = document.getElementById('fileNameResult');
const uploadSection = document.getElementById('upload-section');
const processingSection = document.getElementById('processing-section');
const resultSection = document.getElementById('result-section');
const processButton = document.getElementById('processButton');
const downloadButton = document.getElementById('downloadButton');
const resetButton = document.getElementById('resetButton');
const loadingMessage = document.getElementById('loadingMessage');

// Quando seleciona arquivo
fileInput.addEventListener('change', () => {
  if (fileInput.files.length > 0) {
    const name = fileInput.files[0].name;
    fileName.textContent = `Arquivo selecionado: ${name}`;

    uploadSection.classList.add('hidden');
    processingSection.classList.remove('hidden');

    // estado inicial da tela de processamento
    loadingMessage.textContent = "Aguarde um momento...";
    loadingMessage.style.backgroundColor = "#ffd700";
    loadingMessage.style.color = "black";
    loadingMessage.classList.remove("hidden");

    // esconder resultados anteriores (se houver)
    resultSection.classList.add('hidden');
    downloadButton.classList.add('hidden');
    fileNameResult.textContent = "";
    document.getElementById("analise-margens").innerHTML = "";
    document.getElementById("analise-formatacao").innerHTML = "";
  }
});

// Quando clica em "Processar"
processButton.addEventListener('click', () => {
  if (!fileInput.files.length) return;

  const formData = new FormData();
  formData.append("arquivo", fileInput.files[0]);

  // chamada ao backend
  fetch("http://127.0.0.1:5000/verificar", {
    method: "POST",
    body: formData
  })
  .then(response => {
    if (!response.ok) throw new Error("Resposta do servidor não OK");
    return response.json();
  })
  .then(data => {
    // Exibir resultados mantendo ícones e classes
    // Margens
    const margensEl = document.getElementById("analise-margens");
    if (data.margens_corretas !== undefined) {
      margensEl.innerHTML = data.margens_corretas
        ? `<p class="correto">✅ Margens corretas segundo a ABNT.</p>`
        : `<p class="erro">❌ Margens fora do padrão ABNT.</p>`;
    } else if (data.margens) {
      // compatibilidade com resposta anterior
      margensEl.innerHTML = data.margens.map(l => `<p>${l}</p>`).join("");
    } else {
      margensEl.innerHTML = "";
    }

    // Formatação — cada mensagem já vem com ícone (✅/❌/⚠️)
    const formatEl = document.getElementById("analise-formatacao");
    if (Array.isArray(data.formatacao)) {
      formatEl.innerHTML = data.formatacao.map(msg => {
        let classe = "";
        if (msg.startsWith("✅")) classe = "correto";
        else if (msg.startsWith("❌")) classe = "erro";
        else if (msg.startsWith("⚠️")) classe = "aviso";
        return `<p class="${classe}">${msg}</p>`;
      }).join("");
    } else {
      formatEl.innerHTML = "";
    }

    // Ajustes visuais: manter a caixa de status "Aguarde..." mas alterar texto para concluído
    loadingMessage.textContent = "✅ Verificação concluída!";
    loadingMessage.style.backgroundColor = "#ffd700"; // mantém amarelo
    loadingMessage.style.color = "black";
    loadingMessage.style.fontWeight = "bolder";
    loadingMessage.style.textAlign = "center";
    loadingMessage.style.padding = "12px";
    loadingMessage.style.borderRadius = "6px";
    loadingMessage.classList.remove("hidden");

    // Mostrar nome do arquivo e resultado
    fileNameResult.textContent = fileInput.files[0].name;
    processingSection.classList.add('hidden');
    resultSection.classList.remove('hidden');

    // Mostrar botão de download (ativa para o usuário baixar o arquivo formatado)
    downloadButton.classList.remove('hidden');
  })
  .catch(err => {
    // Erro de rede/servidor
    loadingMessage.textContent = "❌ Erro ao verificar — verifique o backend.";
    loadingMessage.style.backgroundColor = "#ff4d4d";
    loadingMessage.style.color = "white";
    console.error(err);
    processingSection.classList.remove('hidden');
    resultSection.classList.add('hidden');
    downloadButton.classList.add('hidden');
  });
});

// Quando clica em "Baixar já formatado"
downloadButton.addEventListener("click", () => {
  if (!fileInput.files.length) return;

  const formData = new FormData();
  formData.append("arquivo", fileInput.files[0]);

  fetch("http://127.0.0.1:5000/formatar", {
    method: "POST",
    body: formData
  })
  .then(response => {
    if (!response.ok) throw new Error("Erro ao gerar arquivo");
    return response.blob();
  })
  .then(blob => {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = "arquivo_formatado_ABNT.docx";
    document.body.appendChild(a);
    a.click();
    a.remove();
  })
  .catch(err => {
    loadingMessage.textContent = "❌ Erro ao baixar arquivo formatado.";
    loadingMessage.style.backgroundColor = "#ff4d4d";
    loadingMessage.style.color = "white";
    console.error(err);
  });
});

// Quando clica em "Novo Arquivo"
resetButton.addEventListener('click', () => {
  resultSection.classList.add('hidden');
  uploadSection.classList.remove('hidden');
  processingSection.classList.add('hidden');
  fileInput.value = "";
  fileName.textContent = "";
  fileNameResult.textContent = "";
  loadingMessage.textContent = "Aguarde um momento...";
  loadingMessage.style.backgroundColor = "#ffd700";
  loadingMessage.style.color = "black";
  loadingMessage.classList.remove("hidden");
  downloadButton.classList.add('hidden');

  // limpa áreas
  document.getElementById("analise-margens").innerHTML = "";
  document.getElementById("analise-formatacao").innerHTML = "";
});

// Função auxiliar (mantida exatamente, como você pediu)
function formatarTextoABNT(texto) {
  // Remove espaços extras
  texto = texto.trim();

  // Verifica se está todo em maiúsculas (título possível)
  const isUpperCase = texto === texto.toUpperCase();
  // Conta quantidade de palavras
  const wordCount = texto.split(/\s+/).length;

  // Regras:
  // - Títulos geralmente têm poucas palavras (<= 8) e estão em maiúsculo
  // - Se estiver em maiúsculo demais mas for longo, apenas tratamos como seção
  // - Parágrafos têm apenas a primeira letra maiúscula ou são mistos

  if (isUpperCase && wordCount <= 8) {
    // TÍTULO
    return {
      tipo: "titulo",
      texto: texto,
      estilo: {
        negrito: true,
        maiusculo: true,
        tamanhoFonte: 14,
        alinhamento: "centralizado",
        espacamentoAntes: "maior"
      }
    };
  } else if (isUpperCase && wordCount > 8) {
    // SEÇÃO (ex.: introdução, revisão etc), mas toda maiúscula
    return {
      tipo: "secao",
      texto: texto,
      estilo: {
        negrito: true,
        maiusculo: true,
        tamanhoFonte: 12,
        alinhamento: "esquerda",
        recuo: 0
      }
    };
  } else if (/^[A-Z][a-z]/.test(texto)) {
    // Parágrafo normal
    return {
      tipo: "paragrafo",
      texto: texto,
      estilo: {
        negrito: false,
        maiusculo: false,
        tamanhoFonte: 12,
        alinhamento: "justificado",
        primeiraLinhaRecuo: "1.25cm"
      }
    };
  } else {
    // Caso seja citação ou texto especial
    return {
      tipo: "citacao",
      texto: texto,
      estilo: {
        negrito: false,
        italico: true,
        tamanhoFonte: 11,
        alinhamento: "justificado",
        recuoTotal: "4cm"
      }
    };
  }
}
