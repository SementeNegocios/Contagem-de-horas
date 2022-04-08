// @ts-nocheck
///////

//função para puxar todos os eventos dos Sementers na semana passada
function AlocaçõesSemente(){

    //registra o tempo de início da função [ainda não utilizado]
    //var start = new Date().UTC();
    
    //chama a aba Alocações da planilha pelo ID
    var ss = SpreadsheetApp.openById('1wHBOrNyVk6cweSqAvXLQXaoFLSPDMpgzHk_t55NuOF8');
    var sheets = ss.getSheets();                   
    var sheet = ss.getSheetByName('Alocações2');
    //inserir linha depois da linha 3 pra manter as formulas funcionando
    //sheet.insertRowBefore(4);
    
    //chama a aba Salários da planilha pelo ID e puxa os valores de Sementers[0] e Emails[1]
    var reference = ss.getSheetByName('Sementers');
    var sementers = reference.getRange(2,2,reference.getLastRow(),3).getValues();
    
    //define o período a ser consultado, com a data atual, ano e mês atual, sete dias atrás
    var now = new Date();
    var year = now.getFullYear();
    var month = now.getMonth();
    var day = now.getDate();
    
    
    //se quiser rodar o programa até outro dia que não ontem
    var week = day()-7;
    var date = now.toString();   
    //se quiser rodar o programa para o mês inteiro ou para períodos diferentes
    //var week = 1;
    //last.setValue(Date);
    
    /////////////para cada sementer/////////////
    for (var i = 20; i < sementers.length; i++){
    //se quiser rodar começando de outro Sementer
    //for (var i = 32; i < sementers.length; i++){
    ////////////////////////////////////////////
      
      //nome do sementer
      var sementer = sementers[i][0];
      Logger.log(sementer);
      //email do sementer
      var email = sementers[i][1];
      Logger.log(email);
      
      //puxa o calendário do sementer
      //é preciso que o usuário do Google que roda a função tenha se inscrito na agenda a ser consultada
      var calendar = CalendarApp.getCalendarById(email);
      
      //Verifica se a agenda (email) ainda existe e se o usuário que rodou o programa está inscrito na agenda
      if (calendar !== null){
        
        //puxa todos os eventos do período definido anteriormente
        var events = calendar.getEvents(new Date(year, month, week, 00, 00, 00, 00), new Date(year, month, day, 00, 00, 00, 00));
  
        //para cada evento
        for (var j = 0; j < events.length; j++){
          var event = events[j];
  
          //verifica se o evento é de dia inteiro e 
          //[AINDA NÃO IMPLEMENTADO]se foi confirmado pelo sementer && event.email.getGuestStatus()!="NO" && email.getGuestStatus()!="INVITED" && email.getGuestStatus()!="MAYBE"
          if (event.isAllDayEvent()==false) {
          
            //título do evento
            var title = event.getTitle();
  
            //verifica se é um evento da Semente. 
            //Aqui usamos um código de tags entre colchetes [TAG]. Eventos sem tag são considerados eventos pessoais
            if (title.indexOf("[") >= 0 && title.indexOf("]") >= 0){
  
              //[AINDA NÃO IMPLEMENTADO - melhoria de performance]
              //computar o número de eventos com tag
              //criar um range do mesmo tamanho
              //colocar os dados importantes desses eventos no range
              //"pushar" o range inteiro, de uma só vez, para a planilha
  
              //nome do sementer na coluna A
              var c = 1
              var dataRange = sheet.getRange(4,1);
              dataRange.setValue(sementer);
              c++
            
              //dia do início do evento na coluna B
              date = event.getStartTime();
              dataRange = sheet.getRange(4,2);
              dataRange.setValue(date);
              c++
  
              //string entre [] do título do evento na coluna C
              var project = title.substring(title.lastIndexOf("[")+1,title.lastIndexOf("]"));
              var projectup = project.toUpperCase();
              dataRange = sheet.getRange(4,3);
              dataRange.setValue(projectup);
              c++
  
              //calcula a duração do evento, em horas na coluna D
              var duration = (event.getEndTime()-event.getStartTime())/3600000;
              dataRange = sheet.getRange(4,4);
              dataRange.setValue(duration);
              c++
                         //inserir linha depois da linha 3 pra manter as formulas da linha de referência funcionando
              sheet.insertRowBefore(4);
            }
          }
        }
      }
      //caso não encontra a agenda, passa para o próximo Sementer 
      //ainda não dá uma mensagem de erro, ou marca o Sementer com um indicador
      else {
        i++
      }
    }
    
  }
