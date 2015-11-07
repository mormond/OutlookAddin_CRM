(function () {
  'use strict';

  var serviceApiEndpoint = 'https://customerdataservicedemo.azurewebsites.net/api/';
  var customerQuery = 'customer/';
  var orderQuery = 'order/'
  var customer = {};
  var spin8;

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    jQuery(document).ready(function () {
      app.initialize();

      if ($.fn.Pivot) {
        $('.ms-Pivot').Pivot();
      }

      spin8 = fabric.Spinner(jQuery("#spinner-8point")[0]);
      //spin8.start();
      displayItemDetails();
    });
  };

  function customerLookupCallback(result) {
   
    //spin8.stop(); 
         
    if (result.length > 0) {
      customer.lastName = result[0].LastName;
      customer.firstName = result[0].FirstName;
      customer.email = result[0].Email;
      customer.companyName = result[0].CompanyName;
      customer.customerId = result[0].CustomerId;

      $('#company').text(customer.companyName);
      $('#lastName').text(customer.lastName);
      $('#firstName').text(customer.firstName);
      $('#number').text(customer.customerId);

      $("#customer").removeClass("hidden");

      getOrders(customer.customerId);
    }
    else {
      $("#uknownCustomer").removeClass("hidden");
    }
  }

  function getCustomer(email) {

    $.ajax(
      {
        type: "GET",
        dataType: "jsonp",
        url: serviceApiEndpoint + customerQuery + email + "/true/",
        contentType: "text/javascript",
        success: customerLookupCallback
      });
  }

  function makeCell(text) {
    var cell = $('<div class="ms-Grid-col ms-u-sm2 block"></div>');
    var span = $('<span></span>').text(text);
    span.appendTo(cell);

    return cell;
  }

  function orderLookupCallback(result) {
    var orders = result;
    var lastOrder = new Date(1900, 1, 1);

    orders.forEach(function (element) {

      var orderDate = new Date(element.PurchaseDate);
      if (orderDate > lastOrder) {
        lastOrder = orderDate;
      }

      var row = $('<div class="ms-Grid-row"></div>');

      makeCell(element.CustomerId).appendTo(row);
      makeCell(new Date(element.PurchaseDate).toDateString()).appendTo(row);
      makeCell(new Date(element.InvoiceDate).toDateString()).appendTo(row);
      makeCell(element.TotalAmount).appendTo(row);
      $('<div class="ms-Grid-col ms-u-sm4 block"></div>').appendTo(row);

      row.appendTo($('#orders'));

    }, this);

    $("#lastOrder").text(lastOrder.toDateString());
  }

  function getOrders(customerId) {
    $.ajax(
      {
        type: "GET",
        dataType: "jsonp",
        url: serviceApiEndpoint + orderQuery + customerId + "/",
        contentType: "text/javascript",
        success: orderLookupCallback
      });
  }

  function displayItemDetails() {
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item);

    var from;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      from = Office.cast.item.toMessageRead(item).from;
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      from = Office.cast.item.toAppointmentRead(item).organizer;
    }

    if (from) {
      getCustomer(from.emailAddress);
    }
  }
})();
