const xl = require('excel4node');
const moment = require('moment');
moment.locale('es');


let asistencias = require('./jornada.json');

// TODO MODIFICAR LA QUERY PARA QUE TE ENVIE EL CI...
let { nombres, apellidos } = asistencias[0].dispositivo.empleado;

let { nombre } = asistencias[0].dispositivo;

let wb = new xl.Workbook({ dateFormat: 'dd/mm/yyyy hh:mm:ss' });

let ws = wb.addWorksheet('Registro 1');

let style = wb.createStyle({
    font: {
        color: '#000000',
        size: 12,
        bold: true
    },
    alignment: {
        wrapText: true,
        horizontal: 'center',
    },
    fill: {
        type: 'pattern',
        bgColor: '#3c78d8',
        fgColor: '#3c78d8'
    }
});

let normal = wb.createStyle({
    font: {
        size: 10
    }
});

// Titulo
ws.cell(1, 1, 1, 5, true).string('REGISTRO DE JORNADA LABORAL').style(style);
// Nombres y Apellidos:
ws.cell(2, 1, 2, 2, true).string('NOMBRES Y APELLIDOS:');
// Valor
ws.cell(2, 3, 2, 5, true).string(`${nombres} ${apellidos}`).style({ font: { underline: true } });
// CI
ws.cell(3, 1, 3, 2, true).string('#CEDULA:');
// Valor
ws.cell(3, 3).string('1105587388');
// Fecha Reporte
ws.cell(4, 1, 4, 2, true).string('FECHA DE REPORTE:');
let today = new Date;
ws.cell(4, 3, 4, 5, true).date(today);

ws.cell(5, 1).string('#');
ws.cell(5, 2).string('EVENTO');
ws.cell(5, 3).string('HORA');
ws.cell(5, 4).string('DISPOSITIVO');
ws.cell(5, 5).string('UBICACION');



let formated = asistencias.map(a => {
    let objTemp = {
        timestamp: moment(a.hora),
        evento: a.evento.nombre,
        dispositivo: a.dispositivo.nombre,
        ubicacion: [a.latitud, a.longitud]
    };
    return objTemp;
});


let ifilas = 6;
let dias = [];

let sumaHoras = 0;
formated.forEach(rg => {

    let { timestamp: date, evento, dispositivo, ubicacion } = rg;
    let ndate = date.format('dddd DD');
    let nhour = date.format('h:mm:ss');

    if (!dias.includes(ndate)) {
        console.log(`---------------${ndate}--------------`);

        dias.push(ndate);
        ws.cell(ifilas, 1, ifilas, 5, true).string(`${ndate}`);
        ifilas++;

        console.log(` ${evento} - ${ndate} - ${nhour} - ${dispositivo} - ${ubicacion}`);

        ws.cell(ifilas, 1).string('#');
        ws.cell(ifilas, 2).string(`${evento}`);
        ws.cell(ifilas, 3).string(`${nhour}`);
        ws.cell(ifilas, 4).string(`${dispositivo}`);
        ws.cell(ifilas, 5).string(ubicacion.join(','));

    } else {
        console.log(` ${evento} - ${ndate} - ${nhour} - ${dispositivo} - ${ubicacion}`);
        ws.cell(ifilas, 1).string('#');
        ws.cell(ifilas, 2).string(`${evento}`);
        ws.cell(ifilas, 3).string(`${nhour}`);
        ws.cell(ifilas, 4).string(`${dispositivo}`);
        ws.cell(ifilas, 5).string(ubicacion.join(','));
    }
    ifilas++
});


//let lunes = semana.lunes;

/*for (let i = 0; i < lunes.length; i++) {
    const element = lunes[i];
    let text = element.format('LLLL');
    let spt = text.split(' ');
    let hour = spt[spt.length - 1];
    ws.cell(x, y).string(hour);
    x++
}*/



wb.write('hola.xlsx');