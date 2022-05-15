'use strict';

const rp = require('request-promise');
const $ = require('cheerio');
const Excel = require('exceljs');



const url = 'https://www.mwcbarcelona.com/exhibition/2019-exhibitors';
var pageNumber = 1;

const getWebsiteContent = async (url, page, array) => {

  var limitPage = 77;
  page++;
  rp(url)
    .then(function(html) {

      for (let i = 0; i < $('.entity', html).length; i++) {
        array.push($('.entity', html)[i].attribs.href);
      }

      if (page <= limitPage) {
        setTimeout(function() {
          getWebsiteContent("https://www.mwcbarcelona.com/exhibition/2019-exhibitors/page/" + page, page, array);
        }, 1000);

      } else {
        console.log(array);
        getWebsiteDetailContent(array, 0, []);
      }
    })
    .catch(function(err) {
      console.log(err);
    });
}

const getWebsiteDetailContent = async (array, index, finalData) => {


  rp(array[index])
    .then(function(html) {

      var dataContact = {};

      dataContact.name = $('.top-area-container > h2', html).text();
      //
      // console.log($('.mod-content > p', html).length);

      if ($('.mod-content > p', html).length > 4 && typeof $('.mod-content > p', html)[$('.mod-content > p', html).length - 1].children[0] != 'undefined') {
        dataContact.phoneNumber = $('.mod-content > p', html)[$('.mod-content > p', html).length - 1].children[0].data;
      }

      if ($('.email-link', html).length > 0) {
        dataContact.email = $('.email-link', html)[0].attribs.href.split(':')[1];
      }

      if ($('.web-site-link', html).length > 0) {
        dataContact.website = $('.web-site-link', html)[0].attribs.href;
      }

      console.log(dataContact);

      //push data to Array
      finalData.push(dataContact);

      //add index
      index++;

      if (index < array.length) {
        setTimeout(function() {
          getWebsiteDetailContent(array, index, finalData);
        }, 500);

      } else {

        var workbook = new Excel.Workbook();
        var file = 'contacts.xlsx';

        var sheet = workbook.addWorksheet('contacts');

        sheet.columns = [{key:"name", header:"Name"},
        {key: "phoneNumber", header: "Phone Number"},
        {key: "email", header: "Email"},
        {key: "website", header: "Website"}];

        sheet.addRows(finalData);

        workbook.xlsx.writeFile(file)
          .then(function() {
            console.log('Array added and then file saved.')
          }).catch(function(err){
            console.log(err);
          });

      }
      return finalData;
    })
    .catch(function(err) {
      console.log(err);
    });
}


getWebsiteContent(url, pageNumber, []);
