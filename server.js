const express = require('express');
const session = require('express-session');
const XLSX = require('xlsx');
const bodyParser = require('body-parser');
const path = require('path');

const app = express();

// Configuração do middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(session({
   secret: 'seu_segredo_aqui', // Mude para um segredo mais seguro em produção
   resave: false,
   saveUninitialized: true,
}));

// Caminho para os arquivos Excel
const barcodeFilePath = path.join(__dirname, 'Código de barra dos crachás.xlsx');
const hoursFilePath = path.join(__dirname, 'Chronus_014 - BANCO DE HORAS MENSAL POR UNIDADE FUNCIONARIO.xlsx');
const dadosBHFilePath = path.join(__dirname, 'Dados BH.xlsx');

// Ler planilha de códigos de barra
const barcodeWorkbook = XLSX.readFile(barcodeFilePath);
const barcodeSheetName = barcodeWorkbook.SheetNames[0];
const barcodeSheet = XLSX.utils.sheet_to_json(barcodeWorkbook.Sheets[barcodeSheetName]);

// Ler planilha de horas
const hoursWorkbook = XLSX.readFile(hoursFilePath);
const hoursSheetName = hoursWorkbook.SheetNames[0];
const hoursSheet = XLSX.utils.sheet_to_json(hoursWorkbook.Sheets[hoursSheetName]);

// Servir a página de login
app.get('/', (req, res) => {
   res.sendFile(path.join(__dirname, 'login.html'));
});

// Autenticação pelo código de barras
app.post('/login', (req, res) => {
   const barcode = req.body.barcode;

   // Procura na planilha o código de barras
   const user = barcodeSheet.find(row => row['Código de barras'] == barcode);

   if (user) {
      req.session.matricula = user['Chapa'];
      req.session.nome = user['Nome'];
      res.redirect('/home');
   } else {
      res.send('Código de barras inválido, verifique sua autorização. <a href="/">Tente novamente</a>.');
   }
});

// Middleware de autenticação
function isAuthenticated(req, res, next) {
   if (req.session.matricula && req.session.nome) {
      return next();
   }
   res.redirect('/'); // Redirecionar para a página de login se não estiver autenticado
}

// Página após login
app.get('/home', isAuthenticated, (req, res) => {
   const { matricula, nome } = req.session;

   res.send(`
      <!DOCTYPE html>
      <html lang="pt-br">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
          <title>Bem-vindo</title>
      </head>
      <body class="bg-light">
          <div class="container">
              <div class="row justify-content-center" style="margin-top: 100px;">
                  <div class="col-md-6">
                      <div class="card text-center">
                          <div class="card-body">
                              <h1 class="card-title">Bem-vindo ${matricula} - ${nome}</h1>
                              <p class="card-text">Você está agora logado no sistema.</p>
                              
                              <form action="/buscar-nome" method="POST">
                                  <div class="form-group">
                                      <label for="matricula">Digite a matrícula (7 números, começando com '5'):</label>
                                      <input type="text" name="matricula" id="matricula" class="form-control" required pattern="5\\d{6}">
                                  </div>
                                  <button type="submit" class="btn btn-primary">Buscar Nome</button>
                              </form>

                              <p id="resultadoNome" class="mt-3"></p>

                              <a href="/" class="btn btn-danger">Sair</a>
                          </div>
                      </div>
                  </div>
              </div>
              <div class="row">
                  <div class="col text-center">
                      <h5>Usuário logado: ${nome}</h5>
                  </div>
              </div>
          </div>
          
          <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
          <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.9.3/dist/umd/popper.min.js"></script>
          <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
      </body>
      </html>
   `);
});

// Buscar nome da matrícula (apenas acessível após login)
app.post('/buscar-nome', isAuthenticated, (req, res) => {
   const matricula = req.body.matricula;
   const { nome } = req.session;

   // Validação da matrícula
   if (!/^\d{7}$/.test(matricula) || !matricula.startsWith('5')) {
      return res.send(`
         <!DOCTYPE html>
         <html lang="pt-br">
         <head>
             <meta charset="UTF-8">
             <meta name="viewport" content="width=device-width, initial-scale=1.0">
             <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
             <title>Erro</title>
         </head>
         <body class="bg-light">
             <div class="container">
                 <div class="row justify-content-center" style="margin-top: 100px;">
                     <div class="col-md-6">
                         <div class="card text-center">
                             <div class="card-body">
                                 <h1 class="card-title text-danger">Erro de Matrícula</h1>
                                 <p class="card-text">Matrícula inválida. A matrícula deve ter 7 dígitos e começar com '5'.</p>
                                 <a href="/home" class="btn btn-primary">Tentar Novamente</a>
                             </div>
                         </div>
                     </div>
                 </div>
             </div>
             <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
             <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.9.3/dist/umd/popper.min.js"></script>
             <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
         </body>
         </html>
      `);
   }

   const funcionario = hoursSheet.find(row => row['Chapa'] && row['Chapa'].toString() === matricula);
   if (!funcionario) {
      return res.send(`
         <!DOCTYPE html>
         <html lang="pt-br">
         <head>
             <meta charset="UTF-8">
             <meta name="viewport" content="width=device-width, initial-scale=1.0">
             <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
             <title>Erro</title>
         </head>
         <body class="bg-light">
             <div class="container">
                 <div class="row justify-content-center" style="margin-top: 100px;">
                     <div class="col-md-6">
                         <div class="card text-center">
                             <div class="card-body">
                                 <h1 class="card-title text-danger">Matrícula Não Encontrada</h1>
                                 <p class="card-text">A matrícula digitada não foi encontrada. Verifique e tente novamente.</p>
                                 <a href="/home" class="btn btn-primary">Tentar Novamente</a>
                             </div>
                         </div>
                     </div>
                 </div>
             </div>
         </body>
         </html>
      `);
   }

   const nomeFuncionario = funcionario['Nome Funcionario'];
   const saldoFinal = funcionario['Saldo Final Horas'] || "Saldo não encontrado."; 

   res.send(`
      <!DOCTYPE html>
      <html lang="pt-br">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
          <title>Resultado da Pesquisa</title>
      </head>
      <body class="bg-light">
          <div class="container">
              <div class="row justify-content-center" style="margin-top: 100px;">
                  <div class="col-md-6">
                      <div class="card text-center">
                          <div class="card-body">
                              <h5>Usuário logado: ${nome}</h5>
                              <h1 class="card-title">Resultado da Pesquisa</h1>
                              <p class="card-text">Nome do Funcionário: ${nomeFuncionario}</p>
                              <p id="resultadoSaldo" class="mt-3">Saldo do Banco de Horas: ${saldoFinal}</p>

                              <form action="/registrar-bh" method="POST" class="mt-4">
                                  <div class="form-group">
                                      <label for="dataBH">Data do BH:</label>
                                      <input type="date" name="dataBH" id="dataBH" class="form-control" required>
                                  </div>
                                  <div class="form-group">
                                      <label for="horaInicio">Hora de Início:</label>
                                      <input type="time" name="horaInicio" id="horaInicio" class="form-control" required>
                                  </div>
                                  <div class="form-group">
                                      <label for="horaFim">Hora de Fim:</label>
                                      <input type="time" name="horaFim" id="horaFim" class="form-control" required>
                                  </div>
                                  <input type="hidden" name="matricula" value="${matricula}">
                                  <button type="submit" class="btn btn-success">Registrar Banco de Horas</button>
                              </form>

                              <a href="/home" class="btn btn-primary mt-3">Voltar à Página Inicial</a>
                          </div>
                      </div>
                  </div>
              </div>
          </div>
      </body>
      </html>
   `);
});

// Registrar o BH
app.post('/registrar-bh', isAuthenticated, (req, res) => {
   const { matricula, dataBH, horaInicio, horaFim } = req.body;
   const { nome: nomeUsuario } = req.session;

   // Validar horários
   if (horaInicio >= horaFim) {
      return res.send(`
         <!DOCTYPE html>
         <html lang="pt-br">
         <head>
             <meta charset="UTF-8">
             <meta name="viewport" content="width=device-width, initial-scale=1.0">
             <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
             <title>Erro</title>
         </head>
         <body class="bg-light">
             <div class="container">
                 <div class="row justify-content-center" style="margin-top: 100px;">
                     <div class="col-md-6">
                         <div class="card text-center">
                             <div class="card-body">
                                 <h1 class="card-title text-danger">Erro!</h1>
                                 <p class="card-text">O horário de início deve ser menor que o horário de fim.</p>
                                 <a href="/home" class="btn btn-primary">Voltar</a>
                             </div>
                         </div>
                     </div>
                 </div>
             </div>
         </body>
         </html>
      `);
   }

   const dataSolicitacao = new Date().toLocaleDateString('pt-BR');
   const horarioSolicitacao = new Date().toLocaleTimeString('pt-BR');

   // Ler arquivo "Dados BH.xlsx"
   let dadosWorkbook;
   try {
      dadosWorkbook = XLSX.readFile(dadosBHFilePath);
   } catch (error) {
      return res.send('Erro ao ler o arquivo de dados. Verifique se ele existe.');
   }

   const dadosSheetName = dadosWorkbook.SheetNames[0];
   const dadosSheet = dadosWorkbook.Sheets[dadosSheetName];
   const dadosJson = XLSX.utils.sheet_to_json(dadosSheet);

   const funcionario = hoursSheet.find(row => row['Chapa'] && row['Chapa'].toString() === matricula);
   const nomeColaborador = funcionario ? funcionario['Nome Funcionario'] : "Nome não encontrado";
   const saldoFinal = funcionario['Saldo Final Horas'];

   // Verificar se o saldo é negativo
   if (saldoFinal && convertToMinutes(saldoFinal) < 0) {
      return res.send(`
         <!DOCTYPE html>
         <html lang="pt-br">
         <head>
             <meta charset="UTF-8">
             <meta name="viewport" content="width=device-width, initial-scale=1.0">
             <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
             <title>Erro</title>
         </head>
         <body class="bg-light">
             <div class="container">
                 <div class="row justify-content-center" style="margin-top: 100px;">
                     <div class="col-md-6">
                         <div class="card text-center">
                             <div class="card-body">
                                 <h1 class="card-title text-danger">Solicitação falhou!</h1>
                                 <p class="card-text">O colaborador está com o BH negativo. Por favor, procure o Departamento Pessoal.</p>
                                 <a href="/home" class="btn btn-primary">Voltar</a>
                             </div>
                         </div>
                     </div>
                 </div>
             </div>
         </body>
         </html>
      `);
   }

   const novaEntrada = {
      "Matrícula": matricula,
      "Nome": nomeColaborador,
      "Data": dataBH,
      "Horário inicial": horaInicio,
      "Horário final": horaFim,
      "Data solicitação": dataSolicitacao,
      "Horário solicitação": horarioSolicitacao,
      "Usuário": nomeUsuario
   };

   dadosJson.push(novaEntrada);
   const novaSheet = XLSX.utils.json_to_sheet(dadosJson);
   dadosWorkbook.Sheets[dadosSheetName] = novaSheet;

   XLSX.writeFile(dadosWorkbook, dadosBHFilePath);

   res.send(`
      <!DOCTYPE html>
      <html lang="pt-br">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
          <title>Sucesso</title>
      </head>
      <body class="bg-light">
          <div class="container">
              <div class="row justify-content-center" style="margin-top: 100px;">
                  <div class="col-md-6">
                      <div class="card text-center">
                          <div class="card-body">
                              <h1 class="card-title text-success">Sucesso!</h1>
                              <p class="card-text">O registro do BH foi realizado com sucesso!</p>
                              <a href="/home" class="btn btn-primary">Voltar à Página Inicial</a>
                          </div>
                      </div>
                  </div>
              </div>
          </div>
      </body>
      </html>
   `);
});

// Função auxiliar para converter saldo de horas em minutos
function convertToMinutes(saldo) {
   const [hours, minutes] = saldo.split(':').map(Number);
   return hours * 60 + minutes;
}

// Iniciar o servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
   console.log(`Servidor rodando na porta ${PORT}`);
});
