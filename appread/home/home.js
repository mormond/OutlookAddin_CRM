(function(){
  'use strict';
  
  var serviceApiEndpoint = 'https://customerdataservicedemo.azurewebsites.net/api/';
  var customerQuery = 'customer/';
  var orderQuery = 'order/'
  var customer = {};

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();
      
        if ($.fn.Pivot) {
          $('.ms-Pivot').Pivot();
        }

      displayItemDetails();
    });
  };
  
  function customerLookupCallback(result) {
   
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
      
      getOrders(customer.customerId);
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
  
  function makeCell(text)
  {
    var cell = $('<div class="ms-Grid-col ms-u-sm3 block"></div>');
    var span = $('<span></span>').text(text);
    span.appendTo(cell);
    
    return cell;    
  }
  
  function orderLookupCallback(result) {
    var orders = result;

    orders.forEach(function(element) {

    var row = $('<div class="ms-Grid-row"></div>');
    
    makeCell(element.CustomerId).appendTo(row);
    makeCell(element.PurchaseDate).appendTo(row);
    makeCell(element.InvoiceDate).appendTo(row);
    makeCell(element.TotalAmount).appendTo(row);
        
    row.appendTo($('#orders'));
    
    }, this);

    
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


  // Displays the "Subject" and "From" fields, based on the current mail item
  function displayItemDetails(){
    var item = Office.cast.item.toItemRead(Office.context.mailbox.item);

    //jQuery('#subject').text(item.subject);

    var from;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      from = Office.cast.item.toMessageRead(item).from;
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      from = Office.cast.item.toAppointmentRead(item).organizer;
    }

    if (from) {
      
      getCustomer(from.emailAddress);
      
      // jQuery('#from').text(from.displayName);
      // jQuery('#from').click(function(){
      //   app.showNotification(from.displayName, from.emailAddress);
      // });
    }
  }
})();
