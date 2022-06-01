
const ExcelJS = require('exceljs');
const moment = require('moment');
const fs = require('fs');
const { MessageMedia, Buttons } = require('whatsapp-web.js');
const { cleanNumber } = require('./handle')
const DELAY_TIME = 170; //ms
const DIR_MEDIA = `${__dirname}/../mediaSend`;
// import { Low, JSONFile } from 'lowdb'
// import { join } from 'path'
const { saveMessage } = require('../adapter')
/**
 * Enviamos archivos multimedia a nuestro cliente
 * @param {*} number 
 * @param {*} fileName 
 */

const sendMedia = (client, number, fileName) => {
    number = cleanNumber(number)
    const file = `${DIR_MEDIA}/${fileName}`;
    if (fs.existsSync(file)) {
        const media = MessageMedia.fromFilePath(file);
        client.sendMessage(number, media, { sendAudioAsVoice: true });
    }
}

/**
 * Enviamos archivos como notas de voz
 * @param {*} number 
 * @param {*} fileName 
 */

 const sendMediaVoiceNote = (client, number, fileName) => {
    number = cleanNumber(number)
    const file = `${DIR_MEDIA}/${fileName}`;
    if (fs.existsSync(file)) {
        const media = MessageMedia.fromFilePath(file);
        client.sendMessage(number, media ,{ sendAudioAsVoice: true });
    }
}

/**
 * Enviamos un mensaje simple (texto) a nuestro cliente
 * @param {*} number 
 */
const sendMessage = async (client, number = null, text = null, trigger = null) => {
   setTimeout(async () => {
    number = cleanNumber(number)
    const message = text
    client.sendMessage(number, message);
    await readChat(number, message, trigger)
    console.log(`⚡⚡⚡ Enviando mensajes....`);
   },DELAY_TIME)
}

/**
 * Enviamos un mensaje con buttons a nuestro cliente
 * @param {*} number 
 */
const sendMessageButton = async (client, number = null, text = null, actionButtons) => {
    number = cleanNumber(number)
    const { title = null, message = null, footer = null, buttons = [] } = actionButtons;
    let button = new Buttons(message,[...buttons], title, footer);
    client.sendMessage(number, button);

    console.log(`⚡⚡⚡ Enviando mensajes....`);
}


/**
 * Opte
 */
const lastTrigger = (number) => new Promise((resolve, reject) => {
    number = cleanNumber(number)
    const pathExcel = `${__dirname}/../chats/${number}.xlsx`;
    const workbook = new ExcelJS.Workbook();
    if (fs.existsSync(pathExcel)) {
        workbook.xlsx.readFile(pathExcel)
            .then(() => {
                const worksheet = workbook.getWorksheet(1);
                const lastRow = worksheet.lastRow;
                const getRowPrevStep = worksheet.getRow(lastRow.number);
                const lastStep = getRowPrevStep.getCell('C').value;
                resolve(lastStep)
            });
    } else {
        resolve(null)
    }
})

/**
 * Guardar historial de conversacion
 * @param {*} number 
 * @param {*} message 
 */
const readChat = async (number, message, trigger = null) => {
    number = cleanNumber(number);
    await saveMessage( message, trigger, number );
    console.log("Mensaje: "+message);
    console.log("Número: "+number);
    //const flag = /^hello/.test(message);
    console.log('Saved');
}

const sendSolicitudPagoId = async (id, port) => {
    let url = `http://localhost:${port}/whatsapp_web/get_url_solicitud_pago?id=${id}`;
    let response = await fetch(url);

    if (response.ok) { // si el HTTP-status es 200-299, obtener cuerpo de la respuesta (método debajo)
        let json = await response.json();
    } else {
        alert("Error-HTTP: " + response.status);
    }

    console.log(response);
}

module.exports = { sendMessage, sendMedia, lastTrigger, sendMessageButton, readChat, sendMediaVoiceNote,  sendSolicitudPagoId}