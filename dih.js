'use strict';

const rp = require('request-promise');
const $ = require('cheerio');
const Excel = require('exceljs');



const url = 'https://s3platform.jrc.ec.europa.eu/digital-innovation-hubs-tool';
var pageNumber = 1;


const getWebsiteContent = (url, index, array) => {

  rp(url)
    .then(function(html) {
      var data = {};

      for (var i = 0; i < $('.table-data > tr > td ', html).length; i++) {

        if ($(' .table-data > tr > td', html)[i] != null && $(' .table-data > tr > td', html)[i].children[0] != undefined) {


          if ((Math.abs(i / 7 - 5 / 7) % 1 == 0  || Math.abs(i / 7 - 5 / 7) % 1 >0.99) && $('.table-data > tr > td', html)[i].children[0].next != null) {
            data.email = $('.table-data > tr > td', html)[i].children[0].next.children[0].attribs.href.split(':')[1].trim();
          }

           else if (Math.abs(((i - 6) / 7) ) % 1 == 0 && $('.table-data > tr > td', html)[i].children[0].next != null) {
            data.website = $('.table-data > tr > td', html)[i].children[0].next.children[0].attribs.href.trim();
          } else if ((i / 7 - 0 / 7) % 1 == 0) {
            data.name = $('.table-data > tr > td', html)[i].children[0].next.children[0].data;

          } else if ((i / 7 - 1 / 7) % 1 == 0) {
            data.morada = $('.table-data > tr > td', html)[i].children[0].data.trim();
          } else if ((i / 7 - 2 / 7) % 1 == 0) {
            data.pais = $('.table-data > tr > td', html)[i].children[0].data.trim();
          } else if ((i / 7 - 3 / 7) % 1 == 0) {
            data.contactPerson = $('.table-data > tr > td', html)[i].children[0].data.trim();
          } else if ((i / 7 - 4 / 7) % 1 == 0) {
            data.phoneNumber = $('.table-data > tr > td', html)[i].children[0].data.trim();
          }
        }


        if (Math.abs(((i - 6) / 7) % 1) == 0) {
          console.log(data);
          array.push(data);
          data = {};
        }
      }

      var newPage = $('.pager > li > a', html)[2].attribs.href.trim();

      if (index < 62) {
        setTimeout(function() {
          getWebsiteContent(newPage, index + 1, array);
        }, 2000);

      } else if (index ==62 ) {
        var workbook = new Excel.Workbook();
        var file = 'dih.xlsx';

        var sheet = workbook.addWorksheet('contacts');

        sheet.columns = [{
            key: "name",
            header: "DIH Name"
          },
          {
            key: "phoneNumber",
            header: "Phone Number"
          },
          {
            key: "pais",
            header: "País"
          },
          {
            key: "email",
            header: "Email"
          },
          {
            key: "morada",
            header: "Morada"
          },
          {
            key: "website",
            header: "Website"
          },
          {
            key: "contactPerson",
            header: "Pessoa de Contacto"
          }
        ];

        sheet.addRows(array);

        workbook.xlsx.writeFile(file)
          .then(function() {
            console.log('Array added and then file saved.')
          }).catch(function(err) {
            console.log(err);
          });
      }
    })
    .catch(function(err) {
      console.log(err);
    });
}


getWebsiteContent(url, 0, []);
