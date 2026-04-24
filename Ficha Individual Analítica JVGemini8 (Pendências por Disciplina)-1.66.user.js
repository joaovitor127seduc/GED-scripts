// ==UserScript==
// @name         Ficha Individual Analítica JVGemini8 (Pendências por Disciplina)
// @namespace    http://tampermonkey.net/
// @version      1.66
// @description  Ferramentas para analisar a ficha individual do GPE/Sigeduca - Exportação Excel (Com Recomposição em todas as funções)
// @author       Lucas S Monteiro & João Vitor C F & Gemini Refactor
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @match        http://sigeduca.seduc.mt.gov.br/ged/hwgedteladocumento.aspx?0,25
// @icon         data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @license      MIT
// @grant        none
// ==/UserScript==

// --- CONFIGURAÇÕES DE POSIÇÃO ---
const POS_CODIGO = 24;      // Posição do Código/Matrícula do aluno na Tabela 3
const POS_SITUACAO = 14;    // Posição de fallback da Situação na Tabela 5
// -------------------------------------------------------------------------

// Função Inteligente para encontrar a situação independente da posição
function buscarSituacao(tabela) {
    if (!tabela) return "INDEFINIDO";
    let spans = tabela.getElementsByTagName("span");
    for (let i = 0; i < spans.length; i++) {
        let txt = spans[i].textContent.trim().toUpperCase();
        if (["CURSANDO", "TRANSFERIDO DA ESCOLA", "TRANSFERIDO DA TURMA", "APROVADO", "REPROVADO", "DESISTENTE", "MATRICULADO", "ATIVO", "FALECIDO", "REMANEJADO"].includes(txt)) {
            return txt;
        }
    }
    return spans[POS_SITUACAO]?.textContent.trim() || "INDEFINIDO";
}

function parseTable(htmlCode) {
    htmlCode = htmlCode.replaceAll('<span style="font-size: 10px">', "");
    htmlCode = htmlCode.replaceAll('<span style="font-family: Arial">', "");
    htmlCode = htmlCode.replaceAll('</span>', "");
    var tempDiv = document.createElement('div');
    tempDiv.innerHTML = htmlCode;
    var table = tempDiv.querySelector('table');
    var outputMatrix = [];
    var rows = table.querySelectorAll('tr');
    rows.forEach(function (row) {
        var rowValues = [];
        var cells = row.querySelectorAll('td');
        cells.forEach(function (cell) {
            rowValues.push(cell.textContent.trim());
        });
        outputMatrix.push(rowValues);
    });
    return outputMatrix;
}

function pegaCabecalho(dados){
    var cabecalho = []; var calcBimes = 1; var calcOdds = 1;
    var dadosCabec = dados;
    for (let u = 0; u < dadosCabec[0].getElementsByTagName("td").length; u++) {
        if(dadosCabec[0].getElementsByTagName("td")[u].rowSpan == 2){
            cabecalho.push(dadosCabec[0].getElementsByTagName("td")[u].getElementsByTagName("span")[0].textContent.trim());
        }
        if(dadosCabec[0].getElementsByTagName("td")[u].rowSpan == 1 && dadosCabec[0].getElementsByTagName("td")[u].colSpan==2){
            cabecalho.push("N"+calcBimes,"F"+calcBimes);
            calcBimes++;
        }
        if(dadosCabec[0].getElementsByTagName("td")[u].rowSpan == 1 && dadosCabec[0].getElementsByTagName("td")[u].colSpan==1){
            cabecalho.push("NR"+calcOdds);
            calcOdds++;
        }
    }
    return cabecalho;
}

function coletarDados(){
    var infoCabecalho = pegaCabecalho(document.getElementById("content").getElementsByTagName("table")[4].getElementsByTagName("tr"));
    infoCabecalho.shift();

    var output = {'dadosTurma':{},'alunos':{},'materias':[]};
    var tabelas =$('[id=content]');
    var escola = tabelas[0].getElementsByTagName("table")[2].getElementsByTagName("span")[2].textContent;
    var dadosTurma = document.getElementById("content").getElementsByTagName("p")[5].textContent;

    var aluno = document.getElementById("content").getElementsByTagName("table")[3].getElementsByTagName("span")[2]?.textContent.trim() || "SEM_NOME";
    var codigo = document.getElementById("content").getElementsByTagName("table")[3].getElementsByTagName("span")[POS_CODIGO]?.textContent.trim() || "SEM_CODIGO";
    var matricula = document.getElementById("content").getElementsByTagName("table")[3].getElementsByTagName("span")[POS_CODIGO]?.textContent.trim() || "SEM_MATRICULA";
    var faltJust = document.getElementById("content").getElementsByTagName("table")[5].getElementsByTagName("span")[10]?.textContent.trim() || "0";
    var nascimento = document.getElementById("content").getElementsByTagName("table")[3].getElementsByTagName("span")[22]?.textContent.trim() || "";

    var spans = document.getElementById("content").getElementsByTagName("p")[6].getElementsByTagName("span");
    var obs = spans[spans.length - 1]?.textContent.trim() || "";

    output['alunos'][codigo] = {'notas':{},'resultado':'','dadosAluno':{'nome':aluno,'matricula':matricula,'totalFaltas':0,'faltasJust':faltJust,'nascimento':nascimento,'obs':obs}};

    var infos = dadosTurma.split("\n");
    var serie = infos[0].split(">");
    var temp = infos[1].split("TURMA:");
    var turno = temp[0];
    turno = turno.replace("TURNO: ", "");
    temp = temp[1].split("ANO LETIVO:");
    var turma = temp[0];
    var ano = temp[1];

    serie = serie[serie.length - 1];
    turno = turno.trim();
    turma = turma.trim();
    ano = ano.trim();
    serie = serie.trim();

    output["dadosTurma"]["serie"] = serie;
    output["dadosTurma"]["turma"] = turma;
    output["dadosTurma"]["turno"] = turno;
    output["dadosTurma"]["ano"] = ano;
    output["dadosTurma"]["escola"] = escola.split(" - ")[0].trim();
    output["dadosTurma"]["nomeEscola"] = escola.split(" - ")[1].trim();

    var tab5Principal = document.getElementById("content").getElementsByTagName("table")[5];
    output['alunos'][codigo]['resultado'] = buscarSituacao(tab5Principal);

    let result = parseTable(document.getElementById("content").getElementsByTagName("table")[4].outerHTML);

    var compara = 0;
    for (let i = 0; i < result.length; i++) {
        compara = 0;
        let qtdN = result[1].filter(v => String(v).trim() == "N").length;
        let qtdC = result[1].filter(v =>String(v).trim() == "C").length;
        let qtdNC = result[1].filter(v => String(v).trim() == "N/C").length;

        compara= result[0].length + result[1].length - qtdN - qtdC - qtdNC;
        compara = Math.floor(compara);

        if(i==0 || i==1){ continue; }
        if(result[i].length == compara){ result[i].shift(); }

        output.alunos[codigo].notas[result[i][0]] = {};
        if (!output['materias'].includes(result[i][0])) {
            output['materias'].push(result[i][0]);
        }

        for (let j = 1; j < result[i].length; j++) {
            output.alunos[codigo].notas[result[i][0]][infoCabecalho[j]] = result[i][j];
            if(infoCabecalho[j] == "TF"){
                output['alunos'][codigo]['dadosAluno']['totalFaltas'] = output['alunos'][codigo]['dadosAluno']['totalFaltas'] + Number(result[i][j]);
            }
        }
    }

    for (let k = 1; k < tabelas.length; k++) {
        result = parseTable(tabelas[k].getElementsByTagName("table")[4].outerHTML);
        infoCabecalho = pegaCabecalho(tabelas[k].getElementsByTagName("table")[4].getElementsByTagName("tr"));
        infoCabecalho.shift();

        codigo = tabelas[k].getElementsByTagName("table")[3].getElementsByTagName("span")[POS_CODIGO]?.textContent.trim() || "SEM_CODIGO_"+k;
        aluno = tabelas[k].getElementsByTagName("table")[3].getElementsByTagName("span")[2]?.textContent.trim() || "SEM_NOME";
        matricula = tabelas[k].getElementsByTagName("table")[3].getElementsByTagName("span")[POS_CODIGO]?.textContent.trim() || "SEM_MATRICULA";
        faltJust = tabelas[k].getElementsByTagName("table")[5].getElementsByTagName("span")[10]?.textContent.trim() || "0";
        nascimento = tabelas[k].getElementsByTagName("table")[3].getElementsByTagName("span")[22]?.textContent.trim() || "";

        spans = tabelas[k].getElementsByTagName("p")[6].getElementsByTagName("span");
        obs = spans[spans.length - 1]?.textContent.trim() || "";

        if( output['alunos'][codigo] ){}else{
            output['alunos'][codigo] = {'notas':{},'resultado':'','dadosAluno':{'nome':aluno,'matricula':matricula,'totalFaltas':0,'faltasJust':faltJust,'nascimento':nascimento,'obs':obs}};
        }
        compara = 0;
        for (let i = 0; i < result.length; i++) {
            let qtdN = result[1].filter(v => String(v).trim() == "N").length;
            let qtdC = result[1].filter(v =>String(v).trim() == "C").length;
            let qtdNC = result[1].filter(v => String(v).trim() == "N/C").length;

            compara= result[0].length + result[1].length - qtdN - qtdC - qtdNC;
            compara = Math.floor(compara);

            if(i==0 || i==1){ continue; }
            if(result[i].length == compara){ result[i].shift(); }

            var tab5Sec = tabelas[k].getElementsByTagName("table")[5];
            output['alunos'][codigo]['resultado'] = buscarSituacao(tab5Sec);

            output.alunos[codigo].notas[result[i][0]] = {};
            if (!output['materias'].includes(result[i][0])) {
                output['materias'].push(result[i][0]);
            }
            for (let j = 1; j < result[i].length; j++) {
                output.alunos[codigo].notas[result[i][0]][infoCabecalho[j]] = result[i][j];
                if(infoCabecalho[j] == "TF"){
                    output['alunos'][codigo]['dadosAluno']['totalFaltas'] = output['alunos'][codigo]['dadosAluno']['totalFaltas'] + Number(result[i][j]);
                }
            }
        }
    }
    return output;
}

// ---------------- FUNÇÕES DE EXPORTAÇÃO EXCEL ATUALIZADAS ---------------- //

function MapaDeNotas(dataArray, item, titulo) {
    let nomeArquivo = `${titulo} - ${dataArray.dadosTurma.turma} ${dataArray.dadosTurma.turno}.xlsx`;
    let ws_data = [];

    let header = ["Código", "Nome", "Situação"];
    for (let i = 0; i < dataArray.materias.length; i++) {
        header.push(dataArray.materias[i]);
    }
    header.push("Total Faltas Confimadas");
    header.push("Qtde. Notas Vermelhas");
    ws_data.push(header);

    var stringas = ['AVC','INT','BAS','-'];

    Object.keys(dataArray.alunos).forEach(function(cod) {
        let row = [];
        row.push(cod);
        row.push(dataArray.alunos[cod].dadosAluno.nome);
        row.push(dataArray.alunos[cod].resultado);

        let faltaTotal = 0;
        let qtdeVermelha = 0;

        for (var i = 0; i < dataArray.materias.length; i++) {
            var Valor = '';
            if(dataArray.alunos[cod].notas[dataArray.materias[i]] == undefined){
                Valor = '-';
            } else {
                faltaTotal += Number(dataArray.alunos[cod].notas[dataArray.materias[i]].TF || 0);
                if(dataArray.alunos[cod].notas[dataArray.materias[i]][item] == undefined){
                    Valor = '-';
                } else {
                    Valor = dataArray.alunos[cod].notas[dataArray.materias[i]][item];
                }
            }

            if (!stringas.includes(Valor)) {
                let cc = parseFloat(Valor);
                if(!isNaN(cc)){
                    if(cc < 6) qtdeVermelha++;
                    row.push(cc);
                } else {
                    row.push(Valor);
                }
            } else {
                row.push(Valor);
            }
        }

        row.push(faltaTotal - Number(dataArray.alunos[cod].dadosAluno.faltasJust));
        row.push(qtdeVermelha);

        ws_data.push(row);
    });

    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Mapa Notas");
    XLSX.writeFile(wb, nomeArquivo);
}

function MapaDeTodasNotas(dataArray) {
    let nomeArquivo = `Mapa Completo - ${dataArray.dadosTurma.turma} ${dataArray.dadosTurma.turno}.xlsx`;
    let ws_data = [];

    let header = ["Código", "Nome", "Situação", "Período"];
    for (let i = 0; i < dataArray.materias.length; i++) {
        header.push(dataArray.materias[i]);
    }
    header.push("Faltas Acumuladas");
    ws_data.push(header);

    var stringas = ['AVC','INT','BAS','-'];

    const linhasExcel = [
        { chave: 'N1', label: '1º Bimestre', tipo: 'nota' },
        { chave: 'N2', label: '2º Bimestre', tipo: 'nota' },
        { chave: 'NR1', label: '1ª Recomposição', tipo: 'nota' },
        { chave: 'N3', label: '3º Bimestre', tipo: 'nota' },
        { chave: 'N4', label: '4º Bimestre', tipo: 'nota' },
        { chave: 'NR2', label: '2ª Recomposição', tipo: 'nota' },
        { chave: 'Media', label: 'Média Anual', tipo: 'calc' },
        { chave: 'Aprov', label: 'Para Aprovação', tipo: 'calc' }
    ];

    Object.keys(dataArray.alunos).forEach(function(cod) {
        let faltaTotal = 0;
        let mediaano = new Array(dataArray.materias.length).fill(0);

        for (var matIdx = 0; matIdx < dataArray.materias.length; matIdx++) {
             let materiaNotas = dataArray.alunos[cod].notas[dataArray.materias[matIdx]];

             if (materiaNotas) {
                  faltaTotal += Number(materiaNotas.TF || 0);

                  let n1 = parseFloat(materiaNotas['N1']) || 0;
                  let n2 = parseFloat(materiaNotas['N2']) || 0;
                  let n3 = parseFloat(materiaNotas['N3']) || 0;
                  let n4 = parseFloat(materiaNotas['N4']) || 0;
                  let nr1 = parseFloat(materiaNotas['NR1']) || 0;
                  let nr2 = parseFloat(materiaNotas['NR2']) || 0;

                  let finalN1 = n1;
                  let finalN2 = n2;
                  if (n1 < 6.0 && nr1 > n1) { finalN1 = nr1; }
                  if (n2 < 6.0 && nr1 > n2) { finalN2 = nr1; }

                  let finalN3 = n3;
                  let finalN4 = n4;
                  if (n3 < 6.0 && nr2 > n3) { finalN3 = nr2; }
                  if (n4 < 6.0 && nr2 > n4) { finalN4 = nr2; }

                  mediaano[matIdx] = finalN1 + finalN2 + finalN3 + finalN4;
             }
        }
        let faltasReais = faltaTotal - Number(dataArray.alunos[cod].dadosAluno.faltasJust);

        linhasExcel.forEach(linhaCfg => {
            let row = [];
            row.push(cod);
            row.push(dataArray.alunos[cod].dadosAluno.nome);
            row.push(dataArray.alunos[cod].resultado);
            row.push(linhaCfg.label);

            if (linhaCfg.tipo === 'nota') {
                for (var i = 0; i < dataArray.materias.length; i++) {
                    var Valor = '-';
                    if (dataArray.alunos[cod].notas[dataArray.materias[i]]) {
                        let valBruto = dataArray.alunos[cod].notas[dataArray.materias[i]][linhaCfg.chave];
                        if (valBruto != undefined) {
                            Valor = valBruto;
                        }
                    }

                    if (!stringas.includes(Valor) && !isNaN(parseFloat(Valor))) {
                         row.push(parseFloat(Valor));
                    } else {
                         row.push(Valor);
                    }
                }
            } else if (linhaCfg.chave === 'Media') {
                for (var t = 0; t < dataArray.materias.length; t++) {
                    var mediaFinal = (mediaano[t] / 4).toFixed(2);
                    row.push(parseFloat(mediaFinal));
                }
            } else if (linhaCfg.chave === 'Aprov') {
                for (var u = 0; u < dataArray.materias.length; u++) {
                    var mediaApr = (24 - (mediaano[u])).toFixed(2);
                    if (mediaApr >= 24) row.push("-");
                    else if (mediaApr < 0) row.push("APROV");
                    else row.push(parseFloat(mediaApr));
                }
            }
            row.push(faltasReais);
            ws_data.push(row);
        });
        ws_data.push([]);
    });

    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Todas as Notas");
    XLSX.writeFile(wb, nomeArquivo);
}

function MapaParaAprovacao(dataArray) {
    let nomeArquivo = `Mapa Aprovação - ${dataArray.dadosTurma.turma} ${dataArray.dadosTurma.turno}.xlsx`;
    let ws_data = [];

    let header = ["Código", "Nome", "Situação"];
    for (let i = 0; i < dataArray.materias.length; i++) {
        header.push(dataArray.materias[i]);
    }
    ws_data.push(header);

    Object.keys(dataArray.alunos).forEach(function(cod) {
        let row = [];
        let pontosAcumulados = new Array(dataArray.materias.length).fill(0);

        for (var i = 0; i < dataArray.materias.length; i++) {
             let materiaNotas = dataArray.alunos[cod].notas[dataArray.materias[i]];

             if (materiaNotas != undefined) {
                 let n1 = parseFloat(materiaNotas['N1']) || 0;
                 let n2 = parseFloat(materiaNotas['N2']) || 0;
                 let n3 = parseFloat(materiaNotas['N3']) || 0;
                 let n4 = parseFloat(materiaNotas['N4']) || 0;
                 let nr1 = parseFloat(materiaNotas['NR1']) || 0;
                 let nr2 = parseFloat(materiaNotas['NR2']) || 0;

                 let finalN1 = n1;
                 let finalN2 = n2;
                 if (n1 < 6.0 && nr1 > n1) { finalN1 = nr1; }
                 if (n2 < 6.0 && nr1 > n2) { finalN2 = nr1; }

                 let finalN3 = n3;
                 let finalN4 = n4;
                 if (n3 < 6.0 && nr2 > n3) { finalN3 = nr2; }
                 if (n4 < 6.0 && nr2 > n4) { finalN4 = nr2; }

                 pontosAcumulados[i] = finalN1 + finalN2 + finalN3 + finalN4;
             }
        }

        row.push(cod);
        row.push(dataArray.alunos[cod].dadosAluno.nome);
        row.push(dataArray.alunos[cod].resultado);

        for (var u = 0; u < dataArray.materias.length; u++) {
            var mediaApr = (24 - (pontosAcumulados[u])).toFixed(2);
            if (mediaApr >= 24){ row.push("-"); }
            else if (mediaApr <= 0){ row.push("AP"); }
            else { row.push(parseFloat(mediaApr)); }
        }
        ws_data.push(row);
    });

    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Para Aprovação");
    XLSX.writeFile(wb, nomeArquivo);
}

// ---------------- FUNÇÕES DE EXPORTAÇÃO EXCEL ATUALIZADAS ---------------- //

function semNotas(bim, dataArray) {
    let ws_data = [];
    ws_data.push(["Código", "Nome do Aluno", "Turma", "Bimestre", "Disciplina"]);
    let turmaTexto = `${dataArray.dadosTurma.turma} ${dataArray.dadosTurma.turno}`;

    Object.keys(dataArray.alunos).forEach(function(cod) {
        let aluno = dataArray.alunos[cod];

        if (aluno.resultado !== "TRANSFERIDO DA ESCOLA" && aluno.resultado !== "TRANSFERIDO DA TURMA") {
            Object.keys(aluno.notas).forEach(function(materia) {
                for (let i = 1; i <= bim; i++) {
                    let notaKey = 'N' + i;
                    let valorNota = aluno.notas[materia][notaKey];

                    if (valorNota === "-" || valorNota === "" || valorNota === undefined) {
                        ws_data.push([ cod, aluno.dadosAluno.nome, turmaTexto, `${i}º Bimestre`, materia ]);
                    }
                }
            });
        }
    });

    if (ws_data.length === 1) {
        alert("Nenhuma nota faltando encontrada para o período selecionado!");
        return;
    }

    let nomeArquivo = `Relatorio_Sem_Notas_ate_Bimestre_${bim} - ${dataArray.dadosTurma.turma}.xlsx`;
    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Notas Faltantes");
    XLSX.writeFile(wb, nomeArquivo);
}

// --- NOVA FUNÇÃO: Pendências Agrupadas por Disciplina ---
function pendenciasPorDisciplina(bim, dataArray) {
    let ws_data = [];
    ws_data.push(["Disciplina", "Bimestre Pendente", "Código do Aluno", "Nome do Aluno", "Turma"]);
    let turmaTexto = `${dataArray.dadosTurma.turma} ${dataArray.dadosTurma.turno}`;

    // Objeto para agrupar as faltas usando a disciplina como chave principal
    let faltasPorMateria = {};

    Object.keys(dataArray.alunos).forEach(function(cod) {
        let aluno = dataArray.alunos[cod];

        if (aluno.resultado !== "TRANSFERIDO DA ESCOLA" && aluno.resultado !== "TRANSFERIDO DA TURMA") {
            Object.keys(aluno.notas).forEach(function(materia) {
                for (let i = 1; i <= bim; i++) {
                    let notaKey = 'N' + i;
                    let valorNota = aluno.notas[materia][notaKey];

                    if (valorNota === "-" || valorNota === "" || valorNota === undefined) {
                        if (!faltasPorMateria[materia]) {
                            faltasPorMateria[materia] = [];
                        }
                        faltasPorMateria[materia].push({
                            bimestre: `${i}º Bimestre`,
                            cod: cod,
                            nome: aluno.dadosAluno.nome
                        });
                    }
                }
            });
        }
    });

    // Pega todas as disciplinas que têm pendências e organiza em ordem alfabética
    let materiasOrdenadas = Object.keys(faltasPorMateria).sort();

    if (materiasOrdenadas.length === 0) {
        alert("Nenhuma pendência encontrada para o período selecionado!");
        return;
    }

    // Preenche a planilha Excel com as disciplinas agrupadas
    materiasOrdenadas.forEach(materia => {
        let pendencias = faltasPorMateria[materia];

        // Organiza dentro da disciplina por Bimestre e depois por Nome do aluno
        pendencias.sort((a, b) => {
            if (a.bimestre !== b.bimestre) return a.bimestre.localeCompare(b.bimestre);
            return a.nome.localeCompare(b.nome);
        });

        pendencias.forEach(p => {
            ws_data.push([materia, p.bimestre, p.cod, p.nome, turmaTexto]);
        });
    });

    let nomeArquivo = `Pendencias_Por_Disciplina_ate_Bimestre_${bim} - ${dataArray.dadosTurma.turma}.xlsx`;
    var ws = XLSX.utils.aoa_to_sheet(ws_data);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pendências por Disciplina");
    XLSX.writeFile(wb, nomeArquivo);
}
// -------------------------------------------------------------------------- //

function migrarNotas(bim,dataArray){
    var stringOutput = '{';
    var notaB = '';
    Object.keys(dataArray.alunos).forEach(function(key) {
        Object.keys(dataArray.alunos[key].notas).forEach(function(keyM) {
            let vNota = dataArray.alunos[key].notas[keyM]['N'+bim];
            if(vNota !== "-" && vNota !== "" && vNota !== undefined ){
                notaB = vNota.replace(".", ",");
                stringOutput = stringOutput +"'"+keyM+"':'" + notaB +"',";
            }
        });
    });
    stringOutput = stringOutput + "}";
    stringOutput = stringOutput.replace(",}", "}");
    console.log(stringOutput);

    const tempInput = document.createElement("textarea");
    tempInput.value = stringOutput;
    document.body.appendChild(tempInput);
    tempInput.select();
    tempInput.setSelectionRange(0, 99999);
    try {
        document.execCommand("copy");
        alert("Notas Copiadas - Agora cole no lançamento de notas usando outro sript");
    } catch (err) {
      console.error('Erro ao copiar: ', err);
    }
}

function gerarPCA(dataArray,nome){
    var windowFeatures = 'width=${screenWidth},height=${screenHeight}';
    var novaJanela = window.open('', '_blank',windowFeatures);
    var escola = dataArray.dadosTurma.nomeEscola;
    var turma = dataArray.dadosTurma.turma;
    var turno = dataArray.dadosTurma.turno;
    var serie = dataArray.dadosTurma.serie;

    var tabelaHTML = `
  <!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <title>Plano de Recomposição</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; }
    h1, h2, h3, h4 { text-align: center; margin: 4px 0; }
    .titulo { margin-top: 30px; text-align: center; font-weight: bold; font-size: 18px; }
    .subtitulo { text-align: center; font-size: 16px; margin-bottom: 20px; }
    .info { margin: 20px 0; font-size: 13px; }
    table, th, td { border: 1px solid black; border-collapse: collapse; padding: 8px; font-size: 11px; }
    table { width: 100%; margin-bottom: 20px; }
    .assinaturas { display: flex; justify-content: space-around; margin-top: 40px; }
    .assinatura { text-align: center; width: 30%; }
    .assinatura div { border-top: 1px solid #000; margin-top: 60px; }
    .assinatura-centro { text-align: center; width: 100%; margin-top: 60px; }
    .assinatura-centro div { border-top: 1px solid #000; width: 50%; margin: 0 auto; margin-top: 60px; }
    .final { margin-top: 40px; font-size: 14px; text-align: center; }
  </style>
</head>
<body>
  <div style="text-align: center;">
    <img alt="" src="http://sigeduca.seduc.mt.gov.br/geral/imagem/BRASAO.jpg" width="55" height="46">
  </div>
  <h2>Governo de Mato Grosso</h2>
  <h3>SECRETARIA DE ESTADO DE EDUCAÇÃO</h3>
  <h4>SAGR — Secretaria Adjunta de Gestão Regional</h4>
  <h4>SAGE — Secretaria Adjunta de Gestão Educacional</h4>
  <div class="titulo">ANEXO I</div>
  <div class="subtitulo">Plano de Recomposição da Aprendizagem e Compensação de Ausências</div>
  <div class="info">
    Unidade Escolar: <strong>${escola}</strong><br>
    Estudante: <strong>${nome}</strong> &nbsp;&nbsp; ano/série: <strong>${serie} - ${turma}</strong> &nbsp;&nbsp; Turno: <strong>${turno}</strong><br><br>
    Período de compensação da ausência (período que o estudante faltou):<br>
    Data de início: ___/___/______ até a data ___/___/______<br><br>
    Atividades Desenvolvidas:
  </div>
  <table>
    <thead>
      <tr>
        <th>Data</th>
        <th>Componente Curricular</th>
        <th>Habilidade/Objeto de conhecimento em defasagem de aprendizagem diagnosticadas</th>
        <th>Atividade de recomposição realizada</th>
        <th>Observação</th>
      </tr>
    </thead>
    <tbody>
`;

    for (var i = 0; i < dataArray.materias.length; i++) {
        var t = '';
        if (dataArray.materias[i].length > 21) {
            t = dataArray.materias[i].substring(0, 21) + '...';
        } else {
            t = dataArray.materias[i];
        }
        tabelaHTML += `<tr>
        <td></td><td>${t}</td><td></td><td></td><td></td>
      </tr>`;
    }

    tabelaHTML += `</tbody>
  </table>
  <table>
    <tr>
      <th>AVALIAÇÃO (considerando o período de ausências)</th>
    </tr>
    <tr>
      <td>
        (  ) o estudante, neste período, realizou as atividades propostas, porém necessita de continuidade no trabalho com as habilidades indicadas;<br><br>
        (  ) o estudante, neste período, realizou as atividades propostas e consolidou as habilidades;<br><br>
        (  ) o estudante não realizou as atividades propostas.
      </td>
    </tr>
  </table>
  <span><strong>Validação:</strong></span>
  <div class="assinaturas">
    <div class="assinatura"><div></div>Professor Regente <br> Componente Curricular</div>
    <div class="assinatura"><div></div>Professor Regente <br> Componente Curricular</div>
    <div class="assinatura"><div></div>Professor Regente <br> Componente Curricular</div>
  </div>
  <div class="assinatura-centro"><div></div>Coordenador(a) Pedagógico(a)</div>
  <div class="final">
    Justificativa inserida no sistema SigEduca em ___/___/______ <br><br>
    Validado pelo Conselho de Classe em ___/___/______
  </div>
</body>
</html>
`;
    novaJanela.document.write(tabelaHTML);
}

function removeTransf(dataArray){
    var tabelas =$('[id=content]');
    for (let k = 0; k < tabelas.length; k++) {
        let tab5Transf = tabelas[k].getElementsByTagName("table")[5];
        let situacao = buscarSituacao(tab5Transf);

        if(situacao === "TRANSFERIDO DA ESCOLA" || situacao === "TRANSFERIDO DA TURMA"){
            var primeiraDiv = tabelas[k];
            let proximaDiv = primeiraDiv.nextElementSibling;
            if(proximaDiv) proximaDiv.remove();
            let proximoElemento = primeiraDiv.nextElementSibling;
            if(proximoElemento) proximoElemento.remove();
            primeiraDiv.remove();
        }
    }
    alert("Fichas de alunos transferidos de turma e de escola removidos");
}

function gerarLista(dataArray) {
  var windowFeatures = `width=${screen.width},height=${screen.height}`;
  var novaJanela = window.open('', '_blank', windowFeatures);

  var tabelaHTML = `
    <style type="text/css">
      thead { display: table-header-group; }
    </style>
  `;

  tabelaHTML += `<h3 style="text-align:center;">${dataArray.dadosTurma.turma} ${dataArray.dadosTurma.turno}</h3>`;
  tabelaHTML += `<table border="1">
    <thead>
      <tr>
        <th>Código</th><th>Nome</th><th>Situa.</th><th>Nascimento</th><th>Idade</th>
      </tr>
    </thead>
    <tbody>
  `;

  const alunosOrdenados = Object.entries(dataArray.alunos)
    .sort((a, b) => a[1].dadosAluno.nome.localeCompare(b[1].dadosAluno.nome, 'pt', { sensitivity: 'base' }));

  alunosOrdenados.forEach(([cod, aluno]) => {
    tabelaHTML += `<tr>`;
    tabelaHTML += `<td>${cod}</td>`;
    tabelaHTML += `<td>${aluno.dadosAluno.nome}</td>`;
    tabelaHTML += `<td>${aluno.resultado}</td>`;
    tabelaHTML += `<td>${aluno.dadosAluno.nascimento}</td>`;

    let nascimento = aluno.dadosAluno.nascimento.trim();
    let hoje = new Date();
    let [diaNasc, mesNasc, anoNasc] = nascimento.split('/').map(Number);

    let idade = hoje.getFullYear() - anoNasc;
    if (
      hoje.getMonth() + 1 < mesNasc ||
      (hoje.getMonth() + 1 === mesNasc && hoje.getDate() < diaNasc)
    ) {
      idade--;
    }

    tabelaHTML += `<td>${idade}</td>`;
    tabelaHTML += `</tr>`;
  });

  tabelaHTML += `</tbody></table>`;
  novaJanela.document.write(tabelaHTML);
}

(function() {
    'use strict';

    var infos;
    try {
        infos = coletarDados();
        console.log("Dados Coletados com Sucesso!", infos);
    } catch (e) {
        alert("Erro na coleta de dados. Verifique o console.");
        console.error(e);
        return;
    }

    var menuButton = document.createElement('div');
    menuButton.id = 'floatingMenuButton';
    menuButton.innerHTML = 'Analisar';
    document.body.appendChild(menuButton);

    var menuContainer = document.createElement('div');
    menuContainer.id = 'floatingMenuContainer';
    menuContainer.style.display = 'none';
    var subButton;
    var opcaos; var variva;

    // Mapa de notas --------------------------------------------------------------
    var tttt = document.createElement('div');
    tttt.textContent = 'Baixar Excel de Notas ⬇️';
    tttt.style.backgroundColor= '#242420';
    tttt.className = "menu-item";
    tttt.addEventListener('click', function() {
        if(optMapaNotas.style.display == 'none'){
            optMapaNotas.style.display = 'block';
            tttt.textContent = 'Baixar Excel de Notas ⬆️';
        }else{
            optMapaNotas.style.display = 'none';
            tttt.textContent = 'Baixar Excel de Notas ⬇️';
        }
    });
    menuContainer.appendChild(tttt);

    var optMapaNotas = document.createElement('div');
    optMapaNotas.style.display = 'none';

    opcaos = [ "Excel 1º Bimestre", "Excel 2º Bimestre", "Excel 3º Bimestre", "Excel 4º Bimestre", "Excel Notas finais"];
    variva = [ "N1", "N2", "N3", "N4", "MF"];
    for (var i = 0; i < opcaos.length; i++) {
        (function(nome) {
            nome = opcaos[i];var vvv = variva[i];
            subButton = document.createElement('div');
            subButton.className = "menu-item";
            subButton.textContent = nome;
            subButton.addEventListener('click', function() {
                MapaDeNotas(infos,vvv,nome);
            });
            optMapaNotas.appendChild(subButton);
        })(i);
    }

    // Mapa de todas as notas
    subButton = document.createElement('div');
    subButton.className = "menu-item";
    subButton.textContent = 'Excel Notas Completas (Todos Bimestres)';
    subButton.addEventListener('click', function() {
        MapaDeTodasNotas(infos);
    });
    optMapaNotas.appendChild(subButton);

    // Mapa Para Aprovação
    subButton = document.createElement('div');
    subButton.className = "menu-item";
    subButton.textContent = 'Excel Para Aprovação (Cálculo)';
    subButton.addEventListener('click', function() {
        MapaParaAprovacao(infos);
    });
    optMapaNotas.appendChild(subButton);

    menuContainer.appendChild(optMapaNotas);
    menuContainer.appendChild(document.createElement('hr'));

    // Alunos notas (Modelo Original) --------------------------------------------
    var tttt2 = document.createElement('div');
    tttt2.textContent = 'Alunos sem lançamento de notas ⬇️';
    tttt2.style.backgroundColor= '#242420';
    tttt2.className = "menu-item";
    tttt2.addEventListener('click', function() {
        if(optSemNotas.style.display == 'none'){
            optSemNotas.style.display = 'block';
            tttt2.textContent = 'Alunos sem lançamento de notas ⬆️';
        }else{
            optSemNotas.style.display = 'none';
            tttt2.textContent = 'Alunos sem lançamento de notas ⬇️';
        }
    });
    menuContainer.appendChild(tttt2);

    var optSemNotas = document.createElement('div');
    optSemNotas.style.display = 'none';

    opcaos = [ "Excel: Sem notas até 1º Bimestre", "Excel: Sem notas até 2º Bimestre", "Excel: Sem notas até 3º Bimestre", "Excel: Sem notas até 4º Bimestre"];
    variva = [ "1", "2", "3", "4"];
    for (var j = 0; j < opcaos.length; j++) {
        (function(nome) {
            nome = opcaos[j];var vvv = variva[j];
            subButton = document.createElement('div');
            subButton.className = "menu-item";
            subButton.textContent = nome;
            subButton.addEventListener('click', function() {
                semNotas(vvv,infos);
            });
            optSemNotas.appendChild(subButton);
        })(j);
    }
    menuContainer.appendChild(optSemNotas);
    menuContainer.appendChild(document.createElement('hr'));

    // --- NOVA SEÇÃO: Disciplinas com pendências (Agrupado) ---------------------
    var tttt3 = document.createElement('div');
    tttt3.textContent = 'Disciplinas com pendências de notas ⬇️';
    tttt3.style.backgroundColor= '#242420';
    tttt3.className = "menu-item";
    tttt3.addEventListener('click', function() {
        if(optPendDisciplinas.style.display == 'none'){
            optPendDisciplinas.style.display = 'block';
            tttt3.textContent = 'Disciplinas com pendências de notas ⬆️';
        }else{
            optPendDisciplinas.style.display = 'none';
            tttt3.textContent = 'Disciplinas com pendências de notas ⬇️';
        }
    });
    menuContainer.appendChild(tttt3);

    var optPendDisciplinas = document.createElement('div');
    optPendDisciplinas.style.display = 'none';

    var opcaosDisc = [ "Excel: Pendências até 1º Bimestre", "Excel: Pendências até 2º Bimestre", "Excel: Pendências até 3º Bimestre", "Excel: Pendências até 4º Bimestre"];
    var varivaDisc = [ "1", "2", "3", "4"];
    for (var d = 0; d < opcaosDisc.length; d++) {
        (function(nome) {
            nome = opcaosDisc[d]; var vvv = varivaDisc[d];
            subButton = document.createElement('div');
            subButton.className = "menu-item";
            subButton.textContent = nome;
            subButton.addEventListener('click', function() {
                pendenciasPorDisciplina(vvv, infos); // Chama a nova função
            });
            optPendDisciplinas.appendChild(subButton);
        })(d);
    }
    menuContainer.appendChild(optPendDisciplinas);
    menuContainer.appendChild(document.createElement('hr'));
    // --------------------------------------------------------------------------

    var qtdeAlunos = Object.keys(infos.alunos).length;

    if (qtdeAlunos >= 1){
        // opções Secrertario ---------------------------------------------------------
        var tttt23 = document.createElement('div');
        tttt23.textContent = 'Migrar notas ⬇️';
        tttt23.style.backgroundColor= '#242420';
        tttt23.className = "menu-item";
        tttt23.addEventListener('click', function() {
            if(optMigrarNotas.style.display == 'none'){
                optMigrarNotas.style.display = 'block';
                tttt23.textContent = 'Migrar notas ⬇️';
            }else{
                optMigrarNotas.style.display = 'none';
                tttt23.textContent = 'Migrar notas ⬆️';
            }
        });
        menuContainer.appendChild(tttt23);

        var optMigrarNotas = document.createElement('div');
        optMigrarNotas.style.display = 'none';

        opcaos = [ "Copiar 1º Bimestre", "Copiar 2º Bimestre", "Copiar 3º Bimestre", "Copiar 4º Bimestre"];
        variva = [ "1", "2", "3", "4"];
        for (var k = 0; k < opcaos.length; k++) {
            (function(nome) {
                nome = opcaos[k];var vvv = variva[k];
                subButton = document.createElement('div');
                subButton.className = "menu-item";
                subButton.textContent = nome;
                subButton.addEventListener('click', function() {
                    migrarNotas(vvv,infos);
                });
                optMigrarNotas.appendChild(subButton);
            })(k);
        }

        menuContainer.appendChild(optMigrarNotas);
        menuContainer.appendChild(document.createElement('hr'));

        //----- plano de composição
        const primeiraChave = Object.keys(infos.alunos)[0];
        const nome = infos.alunos[primeiraChave]?.dadosAluno.nome || "ALUNO";

        var tttt25 = document.createElement('div');
        tttt25.textContent = 'Gerar Plano de Recomposição ➡️';
        tttt25.style.backgroundColor= '#242420';
        tttt25.className = "menu-item";
        tttt25.addEventListener('click', function() {
            gerarPCA(infos,nome);
        });
        menuContainer.appendChild(tttt25);

        menuContainer.appendChild(document.createElement('hr'));
        var tttt26 = document.createElement('div');
        tttt26.textContent = 'Remover fichas de alunos transferidos ➡️';
        tttt26.style.backgroundColor= '#242420';
        tttt26.className = "menu-item";
        tttt26.addEventListener('click', function() {
            removeTransf(infos);
            menuContainer.style.display = 'none';
        });
        menuContainer.appendChild(tttt26);

        menuContainer.appendChild(document.createElement('hr'));
        var tttt27 = document.createElement('div');
        tttt27.textContent = 'Gerar lista de alunos por situação, idade e transf.';
        tttt27.style.backgroundColor= '#242420';
        tttt27.className = "menu-item";
        tttt27.addEventListener('click', function() {
            gerarLista(infos);
            menuContainer.style.display = 'none';
        });
        menuContainer.appendChild(tttt27);
    }

    document.body.appendChild(menuContainer);

    var style = document.createElement('style');
    style.innerHTML = '@media print { #floatingMenuButton { display: none !important; } }';
    document.head.appendChild(style);

    // Estilos CSS
    var styles = `
        #floatingMenuButton {
            position: fixed;
            top: 10px;
            right: 10px;
            background-color: #27ae60;
            color: white;
            padding: 10px 20px;
            border-radius: 5px;
            cursor: pointer;
            z-index: 1000;
            font-weight: bold;
        }
        #floatingMenuContainer {
            position: fixed;
            top: 40px;
            right: 10px;
            background-color: #ecf0f1;
            border: 1px solid #27ae60;
            border-radius: 5px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            z-index: 1000;
            padding: 10px;
            max-height: 90vh;
            overflow-y: auto;
        }
        .menu-item {
            padding: 8px 10px;
            color: #27ae60;
            cursor: pointer;
            margin-bottom: 2px;
            border-radius: 3px;
        }
        .menu-item:hover {
            background-color: #27ae60;
            color: white;
            cursor: pointer;
        }
    `;

    var styleSheet = document.createElement('style');
    styleSheet.type = 'text/css';
    styleSheet.innerText = styles;
    document.head.appendChild(styleSheet);

    menuButton.addEventListener('click', function() {
        if (menuContainer.style.display === 'none') {
            menuContainer.style.display = 'block';
        } else {
            menuContainer.style.display = 'none';
        }
    });

})();