$(document).ready(function () {

    styleBotoes = `
        font-size:15px; 
        height: 35px; 
        width: 65px; 
        font-weight:bold; 
        cursor:pointer; 
        border-width: 0; 
        margin:15px;   
        padding: 10px 15px;
        background-color: #72c00c;
        color: white; 
    `;



    if (document.getElementById("class_frequency_justification")) {

        var divFrequencia = document.createElement("div");
        divFrequencia.style = "text-align: end;";

        var botaoFrequencia = document.createElement("a");
        botaoFrequencia.innerHTML = "Marcar como Não Registrado";
        botaoFrequencia.id = "botaoFrequencia";
        botaoFrequencia.style = styleBotoes;
        botaoFrequencia.addEventListener("click", function () {
            for (var caixa of document.querySelectorAll('[id*=_status_2]')) { caixa.checked = 'checked' }
            window.scrollTo(0, document.body.scrollHeight);
        });

        divFrequencia.appendChild(botaoFrequencia);
        document.getElementById("class_frequency_justification").parentElement.appendChild(divFrequencia);
    }

    if (document.getElementById("class_rating_course_topic")) {

        // Div do campo de inserção de notas

        var divNotas = document.createElement("div");
        divNotas.style = "text-align: center; margin: 20px;"


        var textArea = document.createElement("textarea");
        textArea.style = "width: 680px; height: 150px;"
        textArea.placeholder = `Coloque aqui a lista das notas que serão processadas de acordo com as instruções abaixo:
- Você pode copiar do excel,n a primeira coluna deve ser o nome e a segunda coluna a nota do aluno.
- Os nomes dos alunos da sua lista devem ser exatamente iguais aos registrados no Saber;
- O modelo seguido pode ser, por exemplo:
        Nome Completo do Aluno 5
        Nome do Outro Aluno 10
        Terceiro Nome do Aluno 7
`;
        divNotas.appendChild(textArea);

        var btnLimparCampo = document.createElement("a");
        btnLimparCampo.innerHTML = "Limpar";
        btnLimparCampo.addEventListener('click', function () {
            textArea.value = "";
        });
        btnLimparCampo.style = styleBotoes;
        divNotas.appendChild(btnLimparCampo)


        var btnProcessarNotas = document.createElement("a");
        btnProcessarNotas.innerHTML = "Processar e Registrar Notas";
        btnProcessarNotas.addEventListener('click', function () {
            copiarOutput(textArea.value);
        });
        btnProcessarNotas.style = styleBotoes;
        divNotas.appendChild(btnProcessarNotas);


        document.getElementById("class_rating_course_topic").parentElement.parentElement.appendChild(divNotas);

        // Div dos botões de coleta de nomes dos alunos

        var divExtras = document.createElement("div");
        divExtras.style = "text-align: center; margin: 20px;"

        var btnGetNomes = document.createElement("a");
        btnGetNomes.innerHTML = "Coletar Lista de Nomes";
        btnGetNomes.addEventListener('click', function () {
            getNomes(textArea);
        });
        btnGetNomes.style = styleBotoes;
        divExtras.appendChild(btnGetNomes);

        var btnGerarExcel = document.createElement("a");
        btnGerarExcel.innerHTML = "Exportar em formato de Excel";
        btnGerarExcel.addEventListener('click', function () {
            gerarExcel();
        });
        btnGerarExcel.style = styleBotoes;
        divExtras.appendChild(btnGerarExcel);

        document.getElementById("class_rating_course_topic").parentElement.parentElement.appendChild(divExtras);

    }

    function copiarOutput(texto) {
        if (texto == "") {
            alert("O texto está vazio. Tente inserir sua lista de notas corretamente de acordo com as instruções.");
        } else {
            tratamento(texto);
        }

    }

    function getNomes(textArea) {
        if (textArea.value != "") {
            if (confirm("Esta ação irá apagar as informações dos alunos colocadas para processar. Deseja continuar?")) {
                copiarNomes();
            }
        } else {
            copiarNomes();
        }
    }

    function copiarNomes() {
        var nomes = "";
        for (var aluno of document.querySelectorAll('[class*=span8]')) {
            if ((aluno.textContent.trim() != "2020 | Governo da Paraíba") && (aluno.textContent.trim() != "Alunos")) {
                nomes = nomes + aluno.textContent.trim() + "\n";
            }
        }

        textArea.value = nomes;
        textArea.select();
        document.execCommand('copy');
        alert("A lista de nomes foi copiada!");
        textArea.value = "";
    }

    function tratamento(texto) { // textArea.value
        texto = texto.split("\n");
        var lista = [];
        zerados = 0;
        texto.forEach(item => {
            lista.push(tratarItem(item));
            if (tratarItem(item)[1] == "0") {
                zerados = zerados + 1;
            }
        });
        processarDados(lista);
    }

    function tratarItem(item) {
        var aux = item.replace("\t", " ").replace("-", " ").replace("\t", " ");
        var nome = "";
        var nota = "0";
        for (var i = 1; i <= aux.length; i++) {
            if (isNumber(item[i - 1])) {
                nome = nome.trim();
                nota = aux.substring(i - 1, aux.length).trim().replace(",", ".");

                break;
            } else {
                nome = aux.substring(0, i);
            }

        }

        return [nome.toLocaleLowerCase(), nota == "0" ? "" : (nota == "10" ? "10.0" : nota)];
    }

    function isNumber(n) {
        return !isNaN(parseFloat(n) && isFinite(n));
    }

    function processarDados(lista) {
        var cont = -1;
        for (var aluno of document.querySelectorAll('[class*=span8]')) {
            lista.forEach(item => {
                if (item[0] == aluno.textContent.trim().toLowerCase()) {
                    var nota = document.getElementById('class_rating_enrollment_ratings_attributes_' + cont + '_rating');
                    nota.value = item[1];
                }
            });
            cont = cont + 1;
        }
        window.scrollTo(0, document.body.scrollHeight);
    }

    function gerarExcel() {

        var lista = []
        var cont = -1;
        var menor_que_7 = 0;
        var igual_a_7 = 0;
        var maior_que_7 = 0;
        for (var aluno of document.querySelectorAll('[class*=span8]')) {
            var item = (aluno.textContent.trim() != "2020 | Governo da Paraíba") && (aluno.textContent.trim() != "Alunos") ?
                aluno.textContent.trim() : "";
            if (document.getElementById('class_rating_enrollment_ratings_attributes_' + cont + '_rating')) {
                var nota = parseFloat(document.getElementById('class_rating_enrollment_ratings_attributes_' + cont + '_rating').value);
                if (nota < 7) menor_que_7++;
                if (nota == 7) igual_a_7++;
                if (nota > 7) maior_que_7++;

                lista.push([item, nota]);
            }
            cont = cont + 1;
        }


        TableToExcel.convert(
            construirTabela(lista, menor_que_7, igual_a_7, maior_que_7, menor_que_7 + igual_a_7 + maior_que_7),
            {
                name: `Relatorio${document.getElementsByClassName("badge course")[0].innerHTML}${document.getElementById("class_rating_day").value.replace("/","_")}.xlsx`,
                sheet: {
                    name: "Relatório"
                }
            });
    }

    function construirTabela(lista, menor_que_7, igual_a_7, maior_que_7, total) {
        var tabela = document.createElement("div");
        tabela.innerHTML = "<table data-cols-width=\"50,20,20,25,20\"></table>";

        tabela.firstChild.insertAdjacentHTML('beforeend', `
        <tr>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            Bimestre
            </td>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            Data da Avaliação
            </td>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            Tipo da Avaliação
            </td>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            Avaliação
            </td>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            Disciplina
            </td>
        </tr>
        <tr>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${document.getElementById("class_rating_bimester")[document.getElementById("class_rating_bimester").selectedIndex].text}
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${document.getElementById("class_rating_day").value}
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${document.getElementById("class_rating_kind")[document.getElementById("class_rating_kind").selectedIndex].text}
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${document.getElementById("class_rating_class_rating_type_id")[document.getElementById("class_rating_class_rating_type_id").selectedIndex].text}
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${document.getElementsByClassName("badge course")[0].innerHTML}
            </td>
        </tr>
        <tr></tr>
        <tr>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            MENOR QUE 7
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${Math.round(menor_que_7 * 100 / total) + "%"}
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${menor_que_7} alunos
            </td>
        </tr>
        <tr>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            IGUAL A 7
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${Math.round(igual_a_7 * 100 / total) + "%"}
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${igual_a_7} alunos
            </td>
        </tr>
        <tr>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            MAIOR QUE 7
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${Math.round(maior_que_7 * 100 / total) + "%"}
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${maior_que_7} alunos
            </td>
        </tr>
        <tr></tr>
        <tr>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            NOME
            </td>
            <td data-f-color=\"FFFFFF\" data-fill-color=\"72C00C\" data-b-a-s=\"thin\" data-f-bold=\"true\" data-a-h=\"center\">
            NOTA
            </td>
        </tr>
        `);

        // Construção a partir da lista
        lista.forEach(item => {
            tabela.firstChild.insertAdjacentHTML('beforeend', `
            <tr>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${item[0]}
            </td>
            <td data-fill-color=\"FFFFFF\" data-a-h=\"center\" data-b-a-s=\"thin\">
            ${item[1] ? item[1] : ""}
            </td>
            </tr>
            `)
        });

        return tabela.firstChild;
    }
});