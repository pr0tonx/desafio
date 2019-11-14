const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');

const credentials = {
  "type": "service_account",
  "project_id": "planilha-258402",
  "private_key_id": "58205151c9e6eb61163665f6b00a3c7647110396",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCjFIbA1G1vfiyt\nUQJBO0DhQeGw17DHuNlEH0gf+TUf/S8wqdK31kge8jvv2EAF4s7oY1zxcze81Q5a\ne8q3XPULxlXIk5m5SAD+7kWnOdLTWFLUUwadEJ5D9qv34Yy1g8FOADrlwJUrGznZ\nlk37Tu0A9q3XHyDU1Od0TQqJV0Xb2W5DIJcrzfWbwR8stLgC9jLyuf5qqfKZAWh0\n9nqWdHyOynuZu3qEaw0I8+RN3XpZILwU3fhQum7zG6VDD5L467AT4aLoYt7uSzdQ\nqbFuBntTLbrbd5zqy8K7WtXeE2f0GRPfrwR8hGNFYZnVZvDZN51zbDTKYKwMBxG4\nMyUR8PONAgMBAAECggEAJLiySDj9RHoLquI/KPuj9iUc4jKLZMmvqLqHheWajA+5\noNJYt9spRcijbPRLpFeoYiU/sEHSxvNNxluyL7xflG1ucxojZxh62uzB4/AuFDeC\n/TsN1e/AR4sDua9A/T2EGWGNYZ9OJ5T7n46MFD73OLyTNAnXHX2seaCAcyEjAqg0\n8eJLjIlMTJY1mHWyrOeAmbtTuNwXHUOTfYB9tiIyReqlup91bSVW6zRcGcJxFOkG\nifdy/DEELk7P7P5RKF7COkE4CFeDYL8FLkAbPz5tfBth3erN7sTd0D0Sa5h7Oyw5\nAtyd2akT+mzHsw8oqpno9O1ijh5DjfsU2uF+zXAo4QKBgQDR4sgc/LxJw5hgLfX7\nk6EWrmPpp6z0ijjxlfL1Vf3QaQk2qN3JoAgfzhps9M9TdywyStXmdHZoExciPntI\nMzA2JljBLeTxvpco1/rDkBQC6G2/uBxXE+VrsvkXbOSO3/qImHH4wlEZ7FYacwpF\nPvqonCbLQNk99nNvQpX81IXKIQKBgQDG6R34UMBtc7whfIRwEeD/87B+A4lv/Ua2\n33/zIwSX25QuTVrRJLtltyxqjDmdPgHjUKA7vkfc44NkGF02IUw3jVghbDpw+FdD\nHtTv9/cZIaVbhzJRdOPx0UaCk+EoAbFdbpSFUJlohWS2kJUQFTQwzPprcyDD2862\ny6sVf0hz7QKBgQC6eaXX80iKtQhFs7AP49tEno5Qg1QsND5hjhs9lDgcmaXA9YmP\n4Oo279QUp/EoNAKFcG3ZAfJNh2CPYToBLNGR2sISaGc3zWDZvgKjC/hrmPwhUT+E\nsj2sUWf0QyBSPHeIMwFXxbVutcbOWxVt7oWflpT1Etmwrq1i1aMS7fMsYQKBgQCI\nn+ClChpSU7d6LMPvEmjAhcrJk3ZYhNiIjdWd1IS4JeuPLjTeCOPrBrksaiq8tbWo\nRF37C0TjFSbPnuiPYKmwUpahRmyR4hJWGRxbw69nBLRGvQMz7h0PoRZUZGy4BQml\nymmbdHQa1d0KhR7OEDJr/q9XFJoBzb4b0qMtveKvNQKBgAwJZJsWbpKmWRjgS8MR\nygCp1/P7mmm2amzH43PhuVxUNComQi84p5Iycxz51r9cylffgV5ikQk+oFWhAqrO\ncTK9ki9F1TrA5WT/CLbJjmKUGtDsXrnPFDAzLaMAWX3gqISaUvw9ahSUue32+nMY\n7DcWsin11rkcSsfURPG+ilgx\n-----END PRIVATE KEY-----\n",
  "client_email": "editor@planilha-258402.iam.gserviceaccount.com",
  "client_id": "115566484963696768947",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/editor%40planilha-258402.iam.gserviceaccount.com"
}

const docId = '13eBuIVARxagXptyY8yzK9kbyqczWS0VCs_5ZMWnLFGE'  /* retirei do código '/edit#gid=0' */

const accessSheet = async() => {
    const doc = new GoogleSpreadsheet(docId);
    await promisify(doc.useServiceAccountAuth)(credentials);
    const info = await promisify(doc.getInfo)();
    const workSheet =  info.worksheets[0];

    // accessing rows information from sheet[0] into 'rows'
    const rows = await promisify(workSheet.getRows)();
    //console.log('Informação da linha: ', rows);

    const presenca = rows[0].engenhariadesoftware.split(': ')[1] * 0.25; // reutilizado da tentativa falha (lá de baixo)
    //console.log('presença: ', presenca);

    const cells = await promisify(workSheet.getCells)({
      'min-row': 4,
      'max-row': 27,
      'return-empty': true,
      'min-col': 3,
      'max-col': 8,
    });
    //console.log(cells);

    // iteranting through each valued cell and filling up a vector called 'vetor'
    for(var i = 4; i <= 27; i++){
      var vetor = [];
        for(var j = 3; j <= 8; j++){       
            for(var k = 0; k < cells.length; k++){
                if(cells[k].row === i && cells[k].col === j){
                  vetor.push(cells[k]);
                  //console.log('informação a respeito do vetor: ', vetor);
                }
            }
        }

        //starting the process which prints back updated values to 'situação' and 'Nota para aprovação final"
        if(vetor[0].value >= presenca){
          vetor[4].value = "Reprovado por Falta";
          vetor[5].value = "0";
        }

        else{
          var media = (Number(vetor[1].value) + Number(vetor[2].value) + Number(vetor[3].value)) / 3;
          //console.log('A média dos alunos é: ', media);

            if(media >= 70){           
              vetor[4].value = "Aprovado";
              vetor[5].value = "0";
            }

            if(media >= 50 && media < 70){
              var naf = 100 - media;  

              vetor[4].value = "Exame Final";
              vetor[5].value = naf.toFixed(2);
            }
            
            if(media < 50){
              vetor[4].value = "Reprovado por Nota";
              vetor[5].value = "0";
            }
        }

  // saving the updated values and priting them back on the current sheet
  vetor[4].save();
  //console.log('Representa a variável situação: ', vetor[4]);
  vetor[5].save();
  //console.log('Representa a variável final: ', vetor[5]);
  }
}
// cathing errors if so
accessSheet().then(() => {
  console.log('finished!');
}).catch(err => {
  console.log('error: ', err);
});

/* uma das tentativas falhas xd. commando row.save() aparenta não printar em linhas que não estejam inicialmente preenchidas
//Respectivas posições da linha para serem chamadas na planilha (usado na tentativa falha)
//Matricula = 'engenhariadesoftware'
//Aluno = '_cokwr'
//Faltas = '_cpzh4'
//P1 = '_cre1l'
//P2 = '_chk2m'
//P3 = '_ciyn3'
//Situação = '_ckd7g'
//Final = '_clrrx'        

for(var row of rows){
          var vetor = [row.engenhariadesoftware,
            row._cokwr,
            row._cpzh4,
            row._cre1l,
            row._chk2m,
            row._ciyn3,
            ,
            ,
            ];
            //console.log(vetor);
        
          var faltas = vetor[2];
          var P1 = vetor[3];
          var P2 = vetor[4];
          var P3 = vetor[5];
          var situacao = vetor[6];
          var final = vetor[7];
        //console.log(faltas, P1, P2, P3)
      
          if(faltas <= presenca){
              var media = (Number(P1) + Number(P2) + Number(P3)) / 3;
              
              if(media >= 70){
                
                vetor[6] = "Aprovado";
                vetor[7] = "0";
             }
              else if(media >= 50 && media < 70){
                var naf = 100 - media;
              
                vetor[6] = "Exame Final";
                vetor[7] = naf.toFixed(2);
              }
              else if(media < 50){
                vetor[6] = "Reprovado por Nota";
                vetor[7] = "0";
               } 
          }
          else{
            vetor[6] = "Reprovado por Falta";
            vetor[7] = "0";
          }
         row._ckd7g = vetor[6];
         row._clrrx = vetor [7];
         row.save();
         //console.log(row);
      } 
*/  
