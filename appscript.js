function checkUrls() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var urls = sheet.getRange('A2:A').getValues(); // Get all URLs from column A starting from row 2
  var statuses = [];
  
  urls.forEach(function(urlRow, index) {
    var url = urlRow[0];
    if (url) {
      var status = checkUrl(url) ? "Active" : "Inactive";
      statuses.push([status]);
    } else {
      statuses.push([""]);
    }
  });

  sheet.getRange(2, 2, statuses.length, 1).setValues(statuses); // Update column B with statuses
}

function checkUrl(url) {
  try {
    // Allow redirects and set a reasonable timeout
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: false,
      escaping: false,
      method: 'get',
      contentType: 'application/json',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
      },
      timeout: 5000
    });

    // Check if the status code is in the range 200-399 or 403, indicating a successful or redirect response or forbidden access
    var statusCode = response.getResponseCode();
    return (statusCode >= 200 && statusCode < 400) || statusCode == 403;
  } catch (e) {
    return false;
  }
}
