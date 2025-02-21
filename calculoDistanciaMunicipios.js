const axios = require('axios');
const xlsx = require('xlsx');
const fs = require('fs');


async function calcularDistancia(origem, destino, chaveApi) {
    const url = 'https://maps.googleapis.com/maps/api/distancematrix/json';
    const params = {
        origins: origem,
        destinations: destino,
        key: chaveApi,
        units: 'metric', 
        language: 'pt-BR' 
    };

    try {
        const resposta = await axios.get(url, { params });
        const dados = resposta.data;

        if (dados.status === 'OK') {
            const distancia = dados.rows[0].elements[0].distance.text;
            const duracao = dados.rows[0].elements[0].duration.text;
            return { distancia, duracao };
        } else {
            console.error(`Erro ao calcular distância entre ${origem} e ${destino}: ${dados.status}`);
            return { distancia: 'Erro', duracao: 'Erro' };
        }
    } catch (erro) {
        console.error(`Erro na requisição: ${erro.message}`);
        return { distancia: 'Erro', duracao: 'Erro' };
    }
}


async function calcularDistanciasEntreListas(listaA, listaB, chaveApi) {
    const resultados = [];

    for (const origem of listaA) {
        for (const destino of listaB) {
            if (origem === destino) {
                resultados.push({ origem, destino, distancia: '0 km', duracao: '0 min' });
            } else {
                const { distancia, duracao } = await calcularDistancia(origem, destino, chaveApi);
                resultados.push({ origem, destino, distancia, duracao });
                console.log(`Calculado: ${origem} -> ${destino} = ${distancia} (${duracao})`);
            }

            await new Promise(resolve => setTimeout(resolve, 1000)); 
        }
    }

    return resultados;
}


function gerarExcel(resultados, listaA, listaB, nomeArquivo) {
    const workbook = xlsx.utils.book_new();

  
    const dadosDistancia = [['Cidade'].concat(listaB.map(cidade => cidade.split(',')[0]))];
    listaA.forEach(origem => {
        const linha = [origem.split(',')[0]];
        listaB.forEach(destino => {
            const resultado = resultados.find(r => r.origem === origem && r.destino === destino);
            linha.push(resultado ? resultado.distancia : 'Erro');
        });
        dadosDistancia.push(linha);
    });
    const sheetDistancia = xlsx.utils.aoa_to_sheet(dadosDistancia);
    xlsx.utils.book_append_sheet(workbook, sheetDistancia, 'Distâncias');

   
    const dadosDuracao = [['Cidade'].concat(listaB.map(cidade => cidade.split(',')[0]))];
    listaA.forEach(origem => {
        const linha = [origem.split(',')[0]];
        listaB.forEach(destino => {
            const resultado = resultados.find(r => r.origem === origem && r.destino === destino);
            linha.push(resultado ? resultado.duracao : 'Erro');
        });
        dadosDuracao.push(linha);
    });
    const sheetDuracao = xlsx.utils.aoa_to_sheet(dadosDuracao);
    xlsx.utils.book_append_sheet(workbook, sheetDuracao, 'Durações');

   
    xlsx.writeFile(workbook, nomeArquivo);
    console.log(`Arquivo Excel gerado com sucesso: ${nomeArquivo}`);
}


async function main() {
    const chaveApi = 'sua chave'; 

    const listaA = [
        'Areia Branca, RN',
        'Baraúna, RN',
        'Grossos, RN',
        'Mossoró, RN',
        'Serra do Mel, RN',
        'Tibau, RN',
        'Apodi, RN',
        'Caraúbas, RN',
        'Felipe Guerra, RN',
        'Governador Dix-Sept Rosado, RN',
        'Campo Grande, RN',
        'Janduís, RN',
        'Messias Targino, RN',
        'Paraú, RN',
        'Triunfo Potiguar, RN',
        'Upanema, RN',
        'Alto do Rodrigues, RN',
        'Assu, RN',
        'Itajá, RN',
        'Jucurutu, RN',
        'Pendências, RN',
        'Porto do Mangue, RN',
        'São Rafael, RN',
        'Água Nova, RN',
        'Coronel João Pessoa, RN',
        'Doutor Severiano, RN',
        'Encanto, RN',
        'Luís Gomes, RN',
        'Major Sales, RN',
        'Riacho de Santana, RN',
        'São Miguel, RN',
        'Venha-Ver, RN',
        'Alexandria, RN',
        'Francisco Dantas, RN',
        'Itaú, RN',
        'José da Penha, RN',
        'Marcelino Vieira, RN',
        'Paraná, RN',
        'Pau dos Ferros, RN',
        'Pilões, RN',
        'Portalegre, RN',
        'Rafael Fernandes, RN',
        'Riacho da Cruz, RN',
        'Rodolfo Fernandes, RN',
        'São Francisco do Oeste, RN',
        'Severiano Melo, RN',
        'Taboleiro Grande, RN',
        'Tenente Ananias, RN',
        'Viçosa, RN',
        'Almino Afonso, RN',
        'Antônio Martins, RN',
        'Frutuoso Gomes, RN',
        'João Dias, RN',
        'Lucrécia, RN',
        'Martins, RN',
        'Olho-dÁgua do Borges, RN',
        'Patu, RN',
        'Rafael Godeiro, RN',
        'Serrinha dos Pintos, RN',
        'Umarizal, RN',
        'Caiçara do Norte, RN',
        'Galinhos, RN',
        'Guamaré, RN',
        'Macau, RN',
        'São Bento do Norte, RN',
        'Afonso Bezerra, RN',
        'Angicos, RN',
        'Caiçara do Rio do Vento, RN',
        'Fernando Pedroza, RN',
        'Jardim de Angicos, RN',
        'Lajes, RN',
        'Pedra Preta, RN',
        'Pedro Avelino, RN',
        'São João do Sabugi, RN',
        'Serra Negra do Norte, RN',
        'Timbaúba dos Batistas, RN',
        'Acari, RN',
        'Carnaúba dos Dantas, RN',
        'Cruzeta, RN',
        'Currais Novos, RN',
        'Equador, RN',
        'Jardim do Seridó, RN',
        'Ouro Branco, RN',
        'Parelhas, RN',
        'Santana do Seridó, RN',
        'São José do Seridó, RN',
        'Bento Fernandes, RN',
        'Jandaíra, RN',
        'João Câmara, RN',
        'Parazinho, RN',
        'Poço Branco, RN',
        'Barcelona, RN',
        'Campo Redondo, RN',
        'Coronel Ezequiel, RN',
        'Jaçanã, RN',
        'Japi, RN',
        'Lagoa de Velhos, RN',
        'Lajes Pintadas, RN',
        'Monte das Gameleiras, RN',
        'Ruy Barbosa, RN',
        'Santa Cruz, RN',
        'São Bento do Trairi, RN',
        'São José do Campestre, RN',
        'São Tomé, RN',
        'Serra de São Bento, RN',
        'Sítio Novo, RN',
        'Tangará, RN',
        'Boa Saúde, RN',
        'Bom Jesus, RN',
        'Brejinho, RN',
        'Ielmo Marinho, RN',
        'Jundiá, RN',
        'Lagoa dAnta, RN',
        'Lagoa de Pedras, RN',
        'Lagoa Salgada, RN',
        'Monte Alegre, RN',
        'Nova Cruz, RN',
        'Passa-e-Fica, RN',
        'Passagem, RN',
        'Riachuelo, RN',
        'Santa Maria, RN',
        'Santo Antônio, RN',
        'São Paulo do Potengi, RN',
        'São Pedro, RN',
        'Senador Elói de Souza, RN',
        'Serra Caiada, RN',
        'Serrinha, RN',
        'Várzea, RN',
        'Vera Cruz, RN',
        'Maxaranguape, RN',
        'Pedra Grande, RN',
        'Pureza, RN',
        'Rio do Fogo, RN',
        'São Miguel do Gostoso, RN',
        'Taipu, RN',
        'Touros, RN',
        'Ceará-Mirim, RN',
        'Macaíba, RN',
        'Nísia Floresta, RN',
        'São Gonçalo do Amarante, RN',
        'São José de Mipibu, RN',
        'Extremoz, RN',
        'Natal, RN',
        'Parnamirim, RN',
        'Arez, RN',
        'Baía Formosa, RN',
        'Canguaretama, RN',
        'Espírito Santo, RN',
        'Goianinha, RN',
        'Montanhas, RN',
        'Pedro Velho, RN',
        'Senador Georgino Avelino, RN',
        'Tibau do Sul, RN',
        'Vila Flor, RN'
    ];

  
    const listaB = [
        'Mossoró, RN',
        'Angicos, RN',
        'Pau dos Ferros, RN',
        'Assú, RN'
    ];

    console.log('Calculando distâncias...');
    const resultados = await calcularDistanciasEntreListas(listaA, listaB, chaveApi);

    
    gerarExcel(resultados, listaA, listaB, 'tabela_distancias.xlsx');
}

main();