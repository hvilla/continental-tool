const fs = require('fs'); 
const csv = require('async-csv');
var excel = require('excel4node');
const path = require('path');

class ContinentalTool{
    
    constructor(folderDir,fileName,headers,withFecha){
        this.withFecha = withFecha;
        this.folderDir = folderDir;
        this.fileName = fileName;
        this.headers = headers;
        this.parsedData = [];
    }
    
    async extractor(){
        var csv_data = fs.readFileSync(`${this.folderDir}/${this.fileName}.csv`)
            .toString() // convert Buffer to string
            .replace('"','')
            .split('\n') // split string to lines
            .map(e => e.trim()) // remove white spaces for each line
            .map(e => e.split(';').map(e => e.trim())); // split each line to array
    
        csv_data.map(data=>{
            let auxData = {};
            this.headers.forEach((name,index) => {
                auxData[name]=data[index];
            });
            this.parsedData.push(auxData);
        });
        return this.parsedData;
    }
    
    async transformData(returnedData){
        var added = [];
        var sanititizedData = [];
    
        returnedData.forEach(data => {
            if(!added.includes(data.Examen) && added.includes(data.Examen) != undefined){
                added.push(data.Examen);
            }
        })
        
        var newHeaders = ["historia","apellidos","nombres","genero","empresa","fecha_creacion","cod_perfil","comentario"];
        added.forEach(data => {
            
            if(this.withFecha){
                newHeaders.push(data,'fecha_muestra','fecha_resultado');
            }else{
                newHeaders.push(data);
            }
        });
    
        added = [];
        
        returnedData.forEach(data => {
            
            if(!added.includes(data.Historia)){
                let row = new Array(newHeaders.length).fill(' ');
                
                row[0] = data.Historia;
                row[1] = data.Apellidos;
                row[2] = data.Nombres;
                row[3] = data.Genero;
                row[4] = data.Empresa;
                row[5] = data.FechaCreacion;
                row[6] = data.CodPerfil;
                row[7] = data.Comentario;
    
                const labIndex = newHeaders.indexOf(data.Examen);
                row[labIndex] = data.Resultado;
                if(this.withFecha){
                    row[labIndex+1] = data.FechaMuestra;
                    row[labIndex+2] = data.FechaResultado;
                }
                
                
                added.push(data.Historia);
                sanititizedData.push(row);
            }else{
                const indexRow = added.indexOf(data.Historia);
    
                let row = sanititizedData[indexRow];
    
                const labIndex = newHeaders.indexOf(data.Examen);
                row[labIndex] = data.Resultado;
                if(this.withFecha){
                    row[labIndex+1] = data.FechaMuestra;
                    row[labIndex+2] = data.FechaResultado;
                }

                sanititizedData[indexRow]=row;
                
            }
        })
        
        sanititizedData.unshift(newHeaders);
        return sanititizedData;
    }
    
    async createFile(data){
        var file = fs.createWriteStream(`${this.folderDir}/output/${this.fileName}.csv`);
        file.on('error', function(err) { });
        data.forEach(line => {
            file.write(line.join(';') + '\n');
        });
        file.end();
        return `${this.fileName} file created`;
    }

    async createFileExcel(data){
        var workbook = new excel.Workbook();

        // Add Worksheets to the workbook
        var worksheet = workbook.addWorksheet('Sheet 1');

        // Create a reusable style
        var style = workbook.createStyle({
        font: {
            color: '#000000',
            size: 12
        }});

        data.map((row,rowIndex) => {
            row.map((cell,columnIndex) => {
                worksheet.cell(rowIndex+1,columnIndex+1).string(`${cell}`).style(style);
                
            });
        });

        workbook.write(path.join(__dirname,`${this.folderDir}/output/${this.fileName}.xlsx`));
        return this.fileName;
    }

    async main(){
        if (!fs.existsSync(`${this.folderDir}/output/`)){
            fs.mkdirSync(`${this.folderDir}/output/`);
        }
        this.extractor().then(parsedData => {
            console.log(`Extraction ${this.fileName} finished now transforming`);
            this.transformData(parsedData).then( dataTransformed => {
                console.log(`Data ${this.fileName} transformed now creating excel files`);
                this.createFileExcel(dataTransformed).then(message => {
                    console.log(message+' parsed to .xls');
                })
            })
        });
    }


}

const folderName = 'Septiembre2019';
const headers=["Historia","Apellidos","Nombres","Genero","Empresa","FechaCreacion","CodPerfil","Comentario","Examen","FechaMuestra","FechaResultado","Resultado"];

const cardiovascular = new ContinentalTool(folderName,'cardiovascular',headers,false);
const gestantes = new ContinentalTool(folderName,'gestantes',headers,false);
const ninos = new ContinentalTool(folderName,'ninos',headers,false);
const primer = new ContinentalTool(folderName,'primer',headers,false);

cardiovascular.main();
gestantes.main();
ninos.main();
primer.main();