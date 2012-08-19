/**
* A simple twitter backup util script. 
* Imporve from https://gist.github.com/3303060  thanks to Johannes Nagl
* @dlqingxi
* You can find a demo by this URL: https://docs.google.com/spreadsheet/ccc?key=0Av3XyC66WqmudEFja18zcGpNblowMXZidTg3Tlg4YkE#gid=0
*/

var CONSUMER_KEY = "=========";
var CONSUMER_SECRET = "===============";


var TZ = "GMT+8";//the timezone 
var BASE = 2;//the tweets would showed from this line

var HOUR_X = "C";
var DAY_X = "B";
var MONTH_X = "A";
    
var colums = {};
colums["month"] = "A";
colums["day"] = "B";
colums["hour"] = "C";
colums["time"] = "D";
colums["author"] = "E";
colums["tweet"] = "F";
colums["operat"] = "G";
colums["tid"] = "H";
       
var colors = {};
colors["00"] = "DarkSalmon";
colors["01"] = "Aqua";
colors["02"] = "Aquamarine";
colors["03"] = "Blue";
colors["04"] = "BlueViolet";
colors["05"] = "Brown";
colors["06"] = "BurlyWood";
colors["07"] = "CadetBlue";
colors["08"] = "Chartreuse";
colors["09"] = "Chocolate";
colors["10"] = "Coral";
colors["11"] = "CornflowerBlue";
colors["12"] = "Crimson";
colors["13"] = "Cyan";
colors["14"] = "DarkBlue";
colors["15"] = "DarkCyan";
colors["16"] = "DarkGoldenRod";
colors["17"] = "DarkGreen";
colors["18"] = "DarkKhaki";
colors["19"] = "DarkMagenta";
colors["20"] = "DarkOliveGreen";
colors["21"] = "Darkorange";
colors["22"] = "DarkOrchid";
colors["23"] = "DarkRed";
colors["24"] = "DarkSalmon";
colors["25"] = "DarkSeaGreen";
colors["26"] = "DarkSlateBlue";
colors["27"] = "DarkTurquoise";
colors["28"] = "DarkViolet";
colors["29"] = "DeepPink";
colors["30"] = "DeepSkyBlue";
colors["31"] = "DodgerBlue";

function getConsumerKey() {
  return CONSUMER_KEY;
}

function getConsumerSecret() {
  return CONSUMER_SECRET;
}

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.addMenu("tweets", [
    { name : "post", functionName : "renderTweetDialog" },
    { name: "get", functionName: "refreshTimeline" }
  ]);
};

function refreshTimeline() {
  authorize();
  
  var requestData = {
    "method": "GET",
    "oAuthServiceName": "twitter",
    "oAuthUseToken": "always"
  };
  
  try { 
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var maxTweetId = sheet.getRange(colums["tid"]  + (BASE-1)).getValue();
    
    var result = UrlFetchApp.fetch(
      "https://api.twitter.com/1/statuses/user_timeline.json?include_entities=true&include_rts=true&count=10&exclude_replies=true" + "&since_id="+maxTweetId,
      requestData);//235682008200790016

    var tweets = Utilities.jsonParse(result.getContentText());
    
    sheet.getRange(colums["tid"]  + (BASE-1)).setValue(tweets[0].id_str);//.clearFormat();
    for (var i = tweets.length - 1; i > -1; i--) {      
      sheet.insertRowAfter(BASE-1);
      sheet.getRange(colums["author"] + BASE).setValue("@" + tweets[i].user.screen_name);
      sheet.getRange(colums["tweet"] + BASE).setValue(tweets[i].text);//.clearFormat();
      sheet.getRange(colums["operat"] + BASE).clearContent();
      sheet.getRange(colums["tid"]  + BASE).setFontColor("white").setValue(tweets[i].id_str);
      
      //sheet.getRange("I" + BASE).setFontColor("black").setValue(tweets[0].id_str);//for debug
      
      var ct = tweets[i].created_at;
      //Sat Aug 18 00:06:29 +0000 2012
      var date = new Date(Date.parse(ct));
      var datestr = Utilities.formatDate(date, TZ, "yyyy/MM/dd HH:mm:ss");
      sheet.getRange(colums["time"] + BASE).setValue(datestr).clearFormat();
      
      var yearstr = datestr.toString().substring(0,4);
      var monstr = datestr.toString().substring(5,7);
      var daystr = datestr.toString().substring(8,10);
      var hourstr = datestr.toString().substring(11,13);
      sheet.getRange(colums["month"] + BASE).setBackground(colors[monstr]);
      sheet.getRange(colums["day"] + BASE).setBackground(colors[daystr]);
      sheet.getRange(colums["hour"] + BASE).setBackground(colors[hourstr]);
      
      
      //color the retweeted items     
      if (tweets[i].favorited) {
        sheet.getRange(colums["tweet"] + BASE).setBackgroundColor("yellow");
      }
      
      if (tweets[i].retweeted) {
        sheet.getRange(colums["tweet"] + BASE).setBackgroundColor("CadetBlue");//this seems doesnt work
      }
      
      if (tweets[i].current_user_retweet) {
        sheet.getRange(colums["tweet"] + BASE).setBackgroundColor("green");
      }   

    }
  }
  catch(e) {
    Logger.log(e); 
  }
}

function onEdit() {
  try {
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    if (sheet.getActiveCell().getColumn() != 7) {
      return;
    }
    
    var rowId = sheet.getActiveCell().getRow();
    var command = sheet.getRange(colums["operat"] + rowId).getValue();
    var tweetId = sheet.getRange(colums["tid"]  + rowId).getValue();
    
    if (command == "") {
      return;
    }
    
    var validCommands = {
      "<3": "fave",
      "rt" : "retweet",
      "reply": "reply"
    };
    
    var app = UiApp.createApplication().setTitle('What do you want to tweet today?');
    
    var handler;
    
    if (command in validCommands) {
      switch(validCommands[command]) {
        case "fave":
          handler = app.createServerHandler("fave");
          break;
        case "retweet":
          handler = app.createServerHandler("retweet");
          break;
        case "reply":
          handler = app.createServerHandler("sendTweet");
          break;
      }
    }
    
    var tweet = app.createTextBox().setName("tweetId").setWidth("100%").setValue(tweetId);
    tweet.setVisible(false);
    handler.addCallbackElement(tweet);
    app.setHeight(100);
    
    var sendButton = app.createButton("Yes", handler);
    
    var dialogPanel = app.createFlowPanel();
    dialogPanel.add(tweet);
 
    if (command == "reply") {
      var tweetText = app.createTextBox().setName("tweet").setWidth("100%").setValue(sheet.getRange(colums["author"] + rowId).getValue());
      dialogPanel.add(tweetText);
      handler.addCallbackElement(tweetText);
    }
    
    dialogPanel.add(sendButton);
    app.add(dialogPanel);
    sheet.show(app);
    
  } catch (e) {
    Logger.log(err);
  }
}

function fave(e) {
  var requestData = {
    "method": "POST",
    "oAuthServiceName": "twitter",
    "oAuthUseToken": "always"
  };
  
  try {
    authorize();
    var result = UrlFetchApp.fetch(
      "https://api.twitter.com/1/favorites/create/" + e.parameter.tweetId + ".json",
      requestData);
  } catch (err) {
    Logger.log(err);
  }

  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function retweet(e) {
  authorize();
  
  var requestData = {
    "method": "POST",
    "oAuthServiceName": "twitter",
    "oAuthUseToken": "always"
  };
  
  try {
    var result = UrlFetchApp.fetch(//e.parameter.tweetId
      "https://api.twitter.com/1/statuses/retweet/" + e.parameter.tweetId + ".json",
      requestData);
  } catch (err) {
    Logger.log(Utilities.jsonStringify(err));
  }

  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function renderTweetDialog() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle("Send Tweet");
  app.setHeight(100);
  
  var helpLabel = app.createLabel("Enter your tweet below:");
  helpLabel.setStyleAttribute("text-align", "justify");
  var tweet = app.createTextBox().setName("tweet").setWidth("100%");
  var sendHandler = app.createServerClickHandler("sendTweet");
  sendHandler.addCallbackElement(tweet);
  var sendButton = app.createButton("Send Tweet", sendHandler);
  
  var dialogPanel = app.createFlowPanel();
  dialogPanel.add(helpLabel);
  dialogPanel.add(tweet);
  dialogPanel.add(sendButton);
  app.add(dialogPanel);
  doc.show(app);
}

function authorize() {
  var oauthConfig = UrlFetchApp.addOAuthService("twitter");
  
  oauthConfig.setAccessTokenUrl(
    "https://api.twitter.com/oauth/access_token");
  oauthConfig.setRequestTokenUrl(
    "https://api.twitter.com/oauth/request_token");
  oauthConfig.setAuthorizationUrl(
    "https://api.twitter.com/oauth/authorize");
  oauthConfig.setConsumerKey(getConsumerKey());
  oauthConfig.setConsumerSecret(getConsumerSecret());
  
  var requestData = {
    "method": "GET",
    "oAuthServiceName": "twitter",
    "oAuthUseToken": "always"
  };
  try {
    var result = UrlFetchApp.fetch(
      "https://api.twitter.com/1/statuses/mentions.json",
      requestData);
  } catch(e) {
    Logger.log(e);
  }
}

function sendTweet(e) {
  var tweet = e.parameter.tweet;
  var tweetId = e.parameter.tweetId;
  
  authorize();
  // Tweet must be URI encoded in order to make it to Twitter safely
  var encodedTweet = encodeURIComponent(tweet);
  var requestData = {
    "method": "POST",
    "oAuthServiceName": "twitter",
    "oAuthUseToken": "always"
  };
  
  if (tweetId) {
    requestData.payload = { "in_reply_to_status_id": tweetId };
  }
  
  try {
    var result = UrlFetchApp.fetch(
      "https://api.twitter.com/1/statuses/update.json?status=" + encodedTweet,
      requestData);
  } catch (e) {
    Logger.log(e);
  }
    
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}
