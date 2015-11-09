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

      spin8 = fabric.Spinner(jQuery("#spinner-8point")[0]);
      //spin8.start();
      
      wireEventHandlers();
      
      /* After initialisation, this is the entry point */
      checkCustomer();
    });
  };

  function checkCustomer() {
    /* Get the current mailbox item */
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item);

    /* If it's an email we need the from property, if an appointment we need organizer */
    var from;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      from = Office.cast.item.toMessageRead(item).from;
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      from = Office.cast.item.toAppointmentRead(item).organizer;
    }

    /* Assuming that's worked, do the customer lookup */
    if (from) {
      getCustomer(from.emailAddress);
    }
  }

  function wireEventHandlers() {
    if ($.fn.Pivot) {
      $('.ms-Pivot').Pivot();
    }

    $("#addCustomer").click(function () { $("#addCustomerDialog").removeClass("hidden"); });
    $("#addCustomerOk").click(function () { $("#addCustomerDialog").addClass("hidden"); });
    $("#addCustomerCancel").click(function () { $("#addCustomerDialog").addClass("hidden"); });
  }

  // Called when the customer lookup returns
  function customerLookupCallback(result) {
   
    //spin8.stop(); 

    // If we get a result, populate our customer record and
    // do a lookup on the custome orders
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

  // Customer lookup - JSONP request to our web service
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

  // Grid row helper function
  function makeCell(text) {
    var cell = $('<div class="ms-Grid-col ms-u-sm2 block"></div>');
    var span = $('<span></span>').text(text);
    span.appendTo(cell);

    return cell;
  }

  // Called when the order lookup returns
  function orderLookupCallback(result) {
    var orders = result;
    var lastOrder = new Date(1900, 1, 1);

    // For each item in the result we need to create a grid row
    // and append it to the orders grid
    // We also need to work out the date of the customer's last
    // order as this is displayed on the customer list item
    orders.forEach(function (element) {

      var orderDate = new Date(element.PurchaseDate);
      if (orderDate > lastOrder) {
        lastOrder = orderDate;
      }

      var newRow = $('<div class="ms-Grid-row"></div>');

      makeCell(element.CustomerId).appendTo(newRow);
      makeCell(new Date(element.PurchaseDate).toDateString()).appendTo(newRow);
      makeCell(new Date(element.InvoiceDate).toDateString()).appendTo(newRow);
      makeCell(element.TotalAmount).appendTo(newRow);
      $('<div class="ms-Grid-col ms-u-sm4 block"></div>').appendTo(newRow);

      newRow.appendTo($('#orders'));

    }, this);

    $("#lastOrder").text(lastOrder.toDateString());
  }

  /* Orders lookup - JSONP request to our web service */
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
})();
