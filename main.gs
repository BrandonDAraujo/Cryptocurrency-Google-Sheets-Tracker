function myFunction() {
  var spread = SpreadsheetApp.getActiveSheet();
  try {
  var limitCounter = 0;
  var response = UrlFetchApp.fetch("https://api.coingecko.com/api/v3/coins/list");
  limitCounter +=1;
  var condensed = response.getContentText();
  var data = JSON.parse(condensed);

  var tickers = [];
  var tZROP = [];
  
  spread.getRange(6, 1).setBackground('orange');
  
  for (z = 0; z <= 100; z++){
    var grabTickers = spread.getRange(7+z, 1).getValues();
    var refineTickers = grabTickers[0][0];
    if (refineTickers != '' && refineTickers != "TZROP"){
      tickers.push({symbol: refineTickers, id:""});
    } 
    else if (refineTickers == "TZROP"){
      tZROP.push(z);
    }
    else if (refineTickers == ''){
      spread.getRange("A"+(z+7)+":P"+(z+7)).setValue('');
    }
  }
  var index = 1;
  for (x = 0; x <= tickers.length-1; x++){
    var filteredObj = data.find(function(item, i){
      if(item.symbol == tickers[x]["symbol"]){
        index = i;
        return i;
      }
    });
    if (tickers[x]["symbol"] == "cvc"){
      tickers[x]["id"] = "civic"
    }else{
      tickers[x]["id"] = data[index]["id"];
    }
  }
  var url = tickers[0]["id"];
  for (a = 1; a <= tickers.length-1; a++){
    url = url + ',' + tickers[a]["id"];
  }
  //Logger.log(tickers);
  
  function links(n, url, tickers){
    var rowOffset = 7;
    var resultsUSD = [];
    var resultsBTC = [];
    var resultsETH = [];
    for (a = 1; a <= n; a++){
      var responseUSD = UrlFetchApp.fetch("https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&ids="+ url +"&order=market_cap_desc&per_page=100&page=" + a +"&sparkline=false");
      var responseBTC = UrlFetchApp.fetch("https://api.coingecko.com/api/v3/coins/markets?vs_currency=btc&ids="+ url +"&order=market_cap_desc&per_page=100&page=" + a +"&sparkline=false");
      var responseETH = UrlFetchApp.fetch("https://api.coingecko.com/api/v3/coins/markets?vs_currency=eth&ids="+ url +"&order=market_cap_desc&per_page=100&page=" + a +"&sparkline=false");
      limitCounter +=3;
      
      var condensedUSD = responseUSD.getContentText();
      var condensedBTC = responseBTC.getContentText();
      var condensedETH = responseETH.getContentText();
      
      var dataUSD = JSON.parse(condensedUSD);
      var dataBTC = JSON.parse(condensedBTC);
      var dataETH = JSON.parse(condensedETH);
      //Logger.log(dataUSD.length);
      
      if (a == 1){
        resultsUSD = resultsUSD.concat(dataUSD);
        resultsBTC = resultsBTC.concat(dataBTC);
        resultsETH = resultsETH.concat(dataETH);
        //Logger.log(JSON.stringify(resultsUSD));
      }else {
      resultsUSD = resultsUSD.concat(dataUSD);
      resultsBTC = resultsBTC.concat(dataBTC);
      resultsETH = resultsETH.concat(dataETH);
      //Logger.log(resultsUSD.length);
      }
    }
    var overallIndex = [];
    for (a = 0; a <= tickers.length-1; a++){
      var filtered = resultsUSD.find(function(item, i){
        if (item.symbol == tickers[a]["symbol"]){
         filteredIndex = i;
          return i;
        } 
      });
      overallIndex.push(filteredIndex);
    }
    var indexLength = overallIndex.length-1 + tZROP.length;
    var n = 0;
    for (a = 0; a <= indexLength; a++){
      if(spread.getRange(rowOffset+a, 1).getValue() == ''){
        while (spread.getRange(rowOffset+a, 1).getValue() == ''){
         rowOffset++;
        }
      }
      if (a+rowOffset-7 == tZROP[0]){
        spread.getRange(rowOffset+a, 9).setFormula('Loading...');
        spread.getRange(rowOffset+a, 14).setFormula('Loading...');
        spread.getRange(rowOffset+a, 9).setFormula('=IMPORTXML("http://securitytokencap.io/currency/tzero", "//h2/text()")');
        spread.getRange(rowOffset+a, 14).setFormula('=IMPORTXML("http://securitytokencap.io/currency/tzero", "//table[@class=\'table table-striped\']/tr[2]/td[2]")');
        n--;
      }else{
      spread.getRange(rowOffset+a,9).setValue(resultsUSD[overallIndex[n]]['current_price']);
      spread.getRange(rowOffset+a,2).setValue(resultsUSD[overallIndex[n]]['name']);
      spread.getRange(rowOffset+a,16).setValue(resultsUSD[overallIndex[n]]['market_cap_rank']);
      
      spread.getRange(rowOffset+a,10).setValue(resultsBTC[overallIndex[n]]['current_price']);
      spread.getRange(rowOffset+a,12).setValue(resultsETH[overallIndex[n]]['current_price']);
      
      var responseVolume = UrlFetchApp.fetch("https://api.coingecko.com/api/v3/coins/"+ resultsUSD[overallIndex[n]]['id'] +"/market_chart?vs_currency=usd&days=1");
        limitCounter +=1;
      var condensedVolume = responseVolume.getContentText();
      var dataVolume = JSON.parse(condensedVolume);
      
      spread.getRange(rowOffset+a,14).setValue(dataVolume['total_volumes'][0][1]);
      spread.getRange(rowOffset+a,15).setValue(resultsUSD[overallIndex[n]]['price_change_percentage_24h']);
      }
      n++;
    }
  }
  
  if (tickers.length > 50){
      links(2, url, tickers);
  }else{
   links(1, url, tickers); 
  }
   spread.getRange(6, 1).setBackground('green');
  Logger.log(limitCounter);
  }
  catch (err){
    spread.getRange(6, 1).setBackground('green');
    Logger.log(limitCounter);
    Logger.log(err);
    
  }
}