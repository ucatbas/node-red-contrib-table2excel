module.exports = function (RED) {
    const Excel = require('exceljs');


    function table2excel(config) {
        RED.nodes.createNode(this, config);
            var node = this;
            node.on('input', function(msg) {
            const workbook = new Excel.Workbook();
            var table_array = msg.array;

            worksheet = workbook.addWorksheet('mysheet');
            worksheet.properties.defaultRowHeight = 120;
            worksheet.columns = [
                { header: 'No',  width: 2.33 },
                { header: 'Bölüm/Takım', width: 8.78  },
                { header: 'Fotoğraf Linki', width: 18  },
                { header: 'Fotoğraf', width: 18  },
                { header: 'Faaliyet', width: 29.56  },
                { header: 'Tehlike', width: 22.67  },
                { header: 'Uygunsuzluk', width: 24.11  },
                { header: 'Uygunsuzluk nerede tespit edildi?\n( İş Kazası, Ramak Kala-Tehlike Bildirimi,\nBakanlık Denetimleri, Müşteri Denetimleri (BSCI Denetimleri), Kuruluş İçi Çapraz İSG Denetimleri, Saha İSG Denetimleri,Yönetici Saha Turları, İSG Ekipmanları Kontrolleri,\nİç Ortam Ölçüm Çalışmaları ) ', width: 30.89  },
                { header: 'Uygunsuzluğu tespit eden kişi/kurum', width: 12.56 },
                { header: 'Olası Kaza Sonucu', width: 14.11  },
                { header: 'Tespit Tarihi', width: 9.89  },
                { header: 'Şiddet', width: 4.11  },
                { header: 'Frekans', width: 4.11  },
                { header: 'Olasılık', width: 4.11  },
                { header: 'Skor', width: 4.11  },
                { header: 'Derece', width: 8.67 },
                { header: 'Önerilen Önlemler', width: 46.56  }
              ];            
            msg.payload = table_array;
            for (var i=0; i<table_array.length; i++)
            {
                var rowValues = [];
                rowValues[1] = (i+1);//no
                rowValues[2] = table_array[i][1];//bolum takım
                rowValues[3] = table_array[i][2];//fotograf
                rowValues[7] = table_array[i][3];//uygunsuzluk-acıklama
                rowValues[8] = table_array[i][4];//konum
                rowValues[9] = table_array[i][5];//kim
                rowValues[11] = table_array[i][6];//tarih
                // 5te de resim adresi olacak

                worksheet.addRow(rowValues);

           }
         
            workbook.xlsx.writeFile('test.xlsx');

            node.send(msg);

        });

    }

    RED.nodes.registerType("table2excel", table2excel);
}
