'use strict';

const rp = require('request-promise');
const $ = require('cheerio');
const Excel = require('exceljs');



const url = 'https://360techindustry.exponor.pt/expositor/lista-de-expositores/';
var pageNumber = 1;

const getWebsiteContent = (url) => {

  var array = [];

  rp(url)
    .then(function(html) {

      for (var i = 0; i < $('.info > p', html).length; i += 5) {
        var data = {};

        if (typeof $('.info > p', html)[i].children[0] != 'undefined')
        data.name = $('.info > p', html)[i].children[0].data;

        if (typeof $('.info > p', html)[i + 2].children[1] != 'undefined')
        data.email  = $('.info > p', html)[i + 2].children[1].attribs.href.split(':')[1];

        if (typeof $('.info > p', html)[i + 3].children[1] != 'undefined')
        data.website = $('.info > p', html)[i + 3].children[1].attribs.href;

        if (typeof $('.info > p', html)[i + 4].children[1] != 'undefined')
        data.phoneNumber = $('.info > p', html)[i + 4].children[1].data;

        array.push(data);
      }

      var workbook = new Excel.Workbook();
      var file = 'contacts360Industry.xlsx';

      var sheet = workbook.addWorksheet('contacts');

      sheet.columns = [{
          key: "name",
          header: "Name"
        },
        {
          key: "phoneNumber",
          header: "Phone Number"
        },
        {
          key: "email",
          header: "Email"
        },
        {
          key: "website",
          header: "Website"
        }
      ];

      sheet.addRows(array);

      workbook.xlsx.writeFile(file)
        .then(function() {
          console.log('Array added and then file saved.')
        }).catch(function(err) {
          console.log(err);
        });

    })
    .catch(function(err) {
      console.log(err);
    });
}

getWebsiteContent(url);
