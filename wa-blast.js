const fs = require('fs');
const http = require('http');
const Excel = require('exceljs');
// const dateFormat = require('dateformat');
const { Client, MessageMedia, Location, List, Buttons, LocalAuth } = require('./index');
const SESSION_FILE_PATH = './session.json';

string_to_array = function(str) {
    return str.trim().split(":");
}

let sessionCfg;
if (fs.existsSync(SESSION_FILE_PATH)) {
    sessionCfg = require(SESSION_FILE_PATH);
}

const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: { headless: false }
});

client.initialize();

client.on('qr', (qr) => {
    // NOTE: This event will not be fired if a session is specified.
    console.log('QR RECEIVED', qr);
});

client.on('authenticated', () => {
    console.log('AUTHENTICATED');
});

client.on('auth_failure', msg => {
    // Fired if session restore was unsuccessfull
    console.error('AUTHENTICATION FAILURE', msg);
});

client.on('ready', () => {
    console.log('READY');

    // Cek User
    let my_number = client.info.me.user
    var options = {
        host: 'munculmotor.com',
        path: '/bot/wa/cek-user.php?hp=' + my_number,
        method: 'GET'
    };
    callback = function(response) {
      var str = '';
      response.on('data', function (chunk) {
        str += chunk;
      });
      response.on('end', function () {
        console.log(str); // User terdaftar
      });
    }
    http.request(options, callback).end();

    // Kirim pesan
    // let numberTes = '6282326222254@c.us'
    // client.sendMessage(numberTes, 'WA sudah siap digunakan');

    // Kirim Gambar
    // const media = MessageMedia.fromFilePath('./pictures/New-Honda-Brio.jpg');
    // client.sendMessage("6282326222254@c.us", media, { caption: "New Honda Brio" })

    // Menggunakan modul exceljs
    //Read a file
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile("./wa-blast.xlsx").then(function () {

        //Get data in worksheet
        var worksheet   = workbook.getWorksheet('data');
        var ws_pesan    = workbook.getWorksheet('pesan');
        var ws_gambar   = workbook.getWorksheet('gambar');

        //Get pesan
        var baris = ws_pesan.getRow(1);
        var pesan = baris.getCell(2).value;

        // Get picture
        var nama_file = ws_gambar.getRow(2).getCell(1);
        // var caption = ws_gambar.getRow(2).getCell(2);
        if (nama_file == "") {
            console.log("Tidak ada gambar")
            var media = "kosong";
        } else {
            try {
            	var gambar = './pictures/' + nama_file;
            	var media = MessageMedia.fromFilePath(gambar);
            } catch {
            	console.log("Gambar tidak ditemukan");
            }
            console.log("Gambar: " + gambar + ", caption: " + pesan)
        } 

        //Get title
        var judul = worksheet.getRow(1);
        var var_nama    = judul.getCell(2);
        var var_a       = judul.getCell(4);
        var var_b       = judul.getCell(5);
        var var_c       = judul.getCell(6);
        var var_d       = judul.getCell(7);
        var var_e       = judul.getCell(8);

        //Get LastRow
        var last_row_number = worksheet.lastRow.number;

        //Looping per row
        //Row and Cell number start from 1
        for(let i=2; i<=last_row_number; i++) {
            let row = worksheet.getRow(i);
            let nama = row.getCell(2).value;
            let hp = row.getCell(3).value;

            // Format hp
	        if (hp.substring(0, 2) == "08") {
	            let no_belakang = hp.substring(2)
	            hp = "628" + no_belakang
	        } else if (hp.substring(0, 3) == "+62") {
	        	hp = hp.substring(1)
	        } else {
	            hp = hp
	        }
	        let olah_hp1 = hp.replace(" ", "");
	        let olah_hp2 = olah_hp1.replace("+","");
	        let olah_hp3 = olah_hp2.replace("-","");
	        let no_hp = olah_hp3
	        let format_hp = no_hp + "@c.us";

            //Variable
            let val_nama    = row.getCell(2).value;
            let val_a       = row.getCell(4).value;
            let val_b       = row.getCell(5).value;
            let val_c       = row.getCell(6).value;
            let val_d       = row.getCell(7).value;
            let val_e       = row.getCell(8).value;
            let olah_pesan1 = pesan.replace("<<" + var_nama +">>", val_nama)
            let olah_pesan2 = olah_pesan1.replace("<<" + var_a +">>", val_a)
            let olah_pesan3 = olah_pesan2.replace("<<" + var_b +">>", val_b)
            let olah_pesan4 = olah_pesan3.replace("<<" + var_c +">>", val_c)
            let olah_pesan5 = olah_pesan4.replace("<<" + var_d +">>", val_d)
            let olah_pesan6 = olah_pesan5.replace("<<" + var_e +">>", val_e)
            let pesan_olah  = olah_pesan6

            //Delay 3.6 detik
            setTimeout(function() {
                client.isRegisteredUser(format_hp).then(function(isRegistered) {
                    var waktu = new Date().toLocaleString("en-US", {timeZone: "Asia/Jakarta"});

                    if(isRegistered) {
                        // Write Success

                        if(media == "kosong") {
                            client.sendMessage(format_hp, pesan_olah);
                            // console.log("Media Tidak Ada");
                        } else {
                            client.sendMessage(format_hp, media, { caption: pesan_olah })
                            // console.log("Media Ada");
                        }
                                    
                        console.log("Data ke-" + (i-1));
                        console.log(no_hp + " - " + nama + " berhasil terkirim");
                        console.log("");

			            // Write result
			            workbook.xlsx.readFile("./hasil.xlsx").then(function() {
			                var ws_berhasil     = workbook.getWorksheet('berhasil');
			                var last_berhasil   = ws_berhasil.lastRow.number;
			                //Add new row
			                let new_row = ws_berhasil.getRow(last_berhasil+1);
			                new_row.getCell(1).value = last_berhasil;
			                new_row.getCell(2).value = waktu;
			                new_row.getCell(3).value = hp;
			                new_row.getCell(4).value = nama;
			                //Save the workbook
			                new_row.commit();
			                workbook.xlsx.writeFile("./hasil.xlsx");
			            }) 

                    } else {
                        // Write Failed
                        workbook.xlsx.readFile("./hasil.xlsx").then(function() {
                            var ws_gagal        = workbook.getWorksheet('gagal');
                            var last_gagal      = ws_gagal.lastRow.number;
                            //Add new row
                            let new_row = ws_gagal.getRow(last_gagal+1);
                            new_row.getCell(1).value = last_gagal;
                            new_row.getCell(2).value = waktu;
                            new_row.getCell(3).value = hp;
                            new_row.getCell(4).value = nama;
                            //Save the workbook
                            new_row.commit();
                            workbook.xlsx.writeFile("./hasil.xlsx");
                        })
                        console.log("Data ke-" + (i-1));
                        console.log(no_hp + " - " + nama + " nomor tidak terdaftar");
                        console.log("");
                    }
                }) 
            }, 3600*i)
            
        }
        // console.log("=== SELESAI ===");
    });

});

client.on('message', async msg => {
    console.log('MESSAGE RECEIVED', msg);

    if (msg.body === '!ping reply') {
        // Send a new message as a reply to the current one
        msg.reply('pong');

    } else if (msg.body === '!ping') {
        // Send a new message to the same chat
        client.sendMessage(msg.from, 'pong');

    } else if (msg.body.startsWith('!sendto ')) {
        // Direct send a new message to specific id
        let number = msg.body.split(' ')[1];
        let messageIndex = msg.body.indexOf(number) + number.length;
        let message = msg.body.slice(messageIndex, msg.body.length);
        number = number.includes('@c.us') ? number : `${number}@c.us`;
        let chat = await msg.getChat();
        chat.sendSeen();
        client.sendMessage(number, message);

    } else if (msg.body.startsWith('!subject ')) {
        // Change the group subject
        let chat = await msg.getChat();
        if (chat.isGroup) {
            let newSubject = msg.body.slice(9);
            chat.setSubject(newSubject);
        } else {
            msg.reply('This command can only be used in a group!');
        }
    } else if (msg.body.startsWith('!echo ')) {
        // Replies with the same message
        msg.reply(msg.body.slice(6));
    } else if (msg.body.startsWith('!desc ')) {
        // Change the group description
        let chat = await msg.getChat();
        if (chat.isGroup) {
            let newDescription = msg.body.slice(6);
            chat.setDescription(newDescription);
        } else {
            msg.reply('This command can only be used in a group!');
        }
    } else if (msg.body === '!leave') {
        // Leave the group
        let chat = await msg.getChat();
        if (chat.isGroup) {
            chat.leave();
        } else {
            msg.reply('This command can only be used in a group!');
        }
    } else if (msg.body.startsWith('!join ')) {
        const inviteCode = msg.body.split(' ')[1];
        try {
            await client.acceptInvite(inviteCode);
            msg.reply('Joined the group!');
        } catch (e) {
            msg.reply('That invite code seems to be invalid.');
        }
    } else if (msg.body === '!groupinfo') {
        let chat = await msg.getChat();
        if (chat.isGroup) {
            msg.reply(`
                *Group Details*
                Name: ${chat.name}
                Description: ${chat.description}
                Created At: ${chat.createdAt.toString()}
                Created By: ${chat.owner.user}
                Participant count: ${chat.participants.length}
            `);
        } else {
            msg.reply('This command can only be used in a group!');
        }
    } else if (msg.body === '!chats') {
        const chats = await client.getChats();
        client.sendMessage(msg.from, `The bot has ${chats.length} chats open.`);
    } else if (msg.body === '!info') {
        let info = client.info;
        client.sendMessage(msg.from, `
            *Connection info*
            User name: ${info.pushname}
            My number: ${info.me.user}
            Platform: ${info.platform}
            WhatsApp version: ${info.phone.wa_version}
        `);
    } else if (msg.body === '!mediainfo' && msg.hasMedia) {
        const attachmentData = await msg.downloadMedia();
        msg.reply(`
            *Media info*
            MimeType: ${attachmentData.mimetype}
            Filename: ${attachmentData.filename}
            Data (length): ${attachmentData.data.length}
        `);
    } else if (msg.body === '!quoteinfo' && msg.hasQuotedMsg) {
        const quotedMsg = await msg.getQuotedMessage();

        quotedMsg.reply(`
            ID: ${quotedMsg.id._serialized}
            Type: ${quotedMsg.type}
            Author: ${quotedMsg.author || quotedMsg.from}
            Timestamp: ${quotedMsg.timestamp}
            Has Media? ${quotedMsg.hasMedia}
        `);
    } else if (msg.body === '!resendmedia' && msg.hasQuotedMsg) {
        const quotedMsg = await msg.getQuotedMessage();
        if (quotedMsg.hasMedia) {
            const attachmentData = await quotedMsg.downloadMedia();
            client.sendMessage(msg.from, attachmentData, { caption: 'Here\'s your requested media.' });
        }
    } else if (msg.body === '!location') {
        msg.reply(new Location(37.422, -122.084, 'Googleplex\nGoogle Headquarters'));
    } else if (msg.location) {
        msg.reply(msg.location);
    } else if (msg.body.startsWith('!status ')) {
        const newStatus = msg.body.split(' ')[1];
        await client.setStatus(newStatus);
        msg.reply(`Status was updated to *${newStatus}*`);
    } else if (msg.body === '!mention') {
        const contact = await msg.getContact();
        const chat = await msg.getChat();
        chat.sendMessage(`Hi @${contact.number}!`, {
            mentions: [contact]
        });
    } else if (msg.body === '!delete') {
        if (msg.hasQuotedMsg) {
            const quotedMsg = await msg.getQuotedMessage();
            if (quotedMsg.fromMe) {
                quotedMsg.delete(true);
            } else {
                msg.reply('I can only delete my own messages');
            }
        }
    } else if (msg.body === '!pin') {
        const chat = await msg.getChat();
        await chat.pin();
    } else if (msg.body === '!archive') {
        const chat = await msg.getChat();
        await chat.archive();
    } else if (msg.body === '!mute') {
        const chat = await msg.getChat();
        // mute the chat for 20 seconds
        const unmuteDate = new Date();
        unmuteDate.setSeconds(unmuteDate.getSeconds() + 20);
        await chat.mute(unmuteDate);
    } else if (msg.body === '!typing') {
        const chat = await msg.getChat();
        // simulates typing in the chat
        chat.sendStateTyping();
    } else if (msg.body === '!recording') {
        const chat = await msg.getChat();
        // simulates recording audio in the chat
        chat.sendStateRecording();
    } else if (msg.body === '!clearstate') {
        const chat = await msg.getChat();
        // stops typing or recording in the chat
        chat.clearState();
    } else if (msg.body === 'jumpto') {
        if (msg.hasQuotedMsg) {
            const quotedMsg = await msg.getQuotedMessage();
            client.interface.openChatWindowAt(quotedMsg.id._serialized);
        }
    }
});

client.on('message_create', (msg) => {
    // Fired on all message creations, including your own
    if (msg.fromMe) {
        // do stuff here
    }
});

client.on('message_revoke_everyone', async (after, before) => {
    // Fired whenever a message is deleted by anyone (including you)
    console.log(after); // message after it was deleted.
    if (before) {
        console.log(before); // message before it was deleted.
    }
});

client.on('message_revoke_me', async (msg) => {
    // Fired whenever a message is only deleted in your own view.
    console.log(msg.body); // message before it was deleted.
});

client.on('message_ack', (msg, ack) => {
    /*
        == ACK VALUES ==
        ACK_ERROR: -1
        ACK_PENDING: 0
        ACK_SERVER: 1
        ACK_DEVICE: 2
        ACK_READ: 3
        ACK_PLAYED: 4
    */

    if(ack == 3) {
        // The message was read
    }
});

client.on('group_join', (notification) => {
    // User has joined or been added to the group.
    console.log('join', notification);
    notification.reply('User joined.');
});

client.on('group_leave', (notification) => {
    // User has left or been kicked from the group.
    console.log('leave', notification);
    notification.reply('User left.');
});

client.on('group_update', (notification) => {
    // Group picture, subject or description has been updated.
    console.log('update', notification);
});

client.on('change_battery', (batteryInfo) => {
    // Battery percentage for attached device has changed
    const { battery, plugged } = batteryInfo;
    console.log(`Battery: ${battery}% - Charging? ${plugged}`);
});

client.on('disconnected', (reason) => {
    console.log('Client was logged out', reason);
});

