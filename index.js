const XLSX = require("xlsx");
const wbm = require('wbm');
const prompt = require('prompt-sync')();

console.log('WhatsApp Sender de PEDIDOS')

console.log('Procesando LIBRO DE TRABAJO de PEDIDOS...')
const workbook = XLSX.readFile("Pedidos.xlsm");

console.log('Plantillas disponibles :', workbook.SheetNames);
const worksheet = workbook.Sheets['Totales'];
console.log('Procesando la plantilla `Totales`...');
const rows = XLSX.utils.sheet_to_json(worksheet);
var bulkList = [];

console.log('Armando al lista Masiva...');
rows.forEach(async iterator => { 
    console.log(iterator);

iterator.Nombre + '\n';
    let msg = '¡Buenas tardes ' + iterator.Nombre + '!\n'

    if(iterator.Delivery === 'Sí' || iterator.Delivery === 'Si' || iterator.Delivery === 'si') {
        msg += 
        '¡Tu pedido está listo! El jueves entre las 9 y las 12 hs aproximadamente estaremos '+
        'visitando tu domicilio para entregártelo.\n'+    
        'El total es $ ' + iterator.Total + '. Agradeceremos si podés abonar con cambio.\n';
    } else {
        msg += 
        '¡Tu pedido está listo! Te esperamos en nuestra sede de calle 25 de Mayo número 66, \n'+
        'el jueves entre las 10 y las 13:30 para retirarlo.\n'+
        'El total es $ ' + iterator.Total + '. Agradeceremos si podés abonar con cambio.\n'+
        'Te recordamos que si tenés alguna caja para darnos, nos será muy útil para futuros pedidos.';
    };
    msg += '¡Gracias por tu compra!\n'
    msg += 'Pedido Numero ' + iterator["Núm. Pedido"] + '\n';

    /* CRAPPY JSON transformation */
    let crappyJSON = iterator["Código JSON"];
    position = crappyJSON.search("items");
    let items = String(crappyJSON).substring(position);
    items = items.substring(0,items.length-1-1);
    const replacer = new RegExp("%0A", 'g');
    items = items.replace(replacer, '\n',);
    items = '\n' + items;
    msg += items;
    
    var node = {
        phone: iterator.Teléfono,
        message : msg
    }
    bulkList.push(node);
});
console.log(bulkList);
console.log('Cantidad de mensajes a enviar: ', bulkList.length);


/* Whatsapp BOT */
console.log('Arrancando BOT de Whatsapp...');
wbm.start( {showBrowser: false, session: true} ).then(async () => {

    prompt('ENTER para continuar y empezar envios BOT...');

    for( inx in bulkList ) { 
        console.log(bulkList[inx].phone, bulkList[inx].message);
        await wbm.sendTo(bulkList[inx].phone, bulkList[inx].message);
    };
    await wbm.end();
    
    prompt('ENTER para finalizar...');
}).catch(err => console.log(err));


