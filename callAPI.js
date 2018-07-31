function callTextAnalyticsAPI(type, description) {
	debugger;	
	var params = {
		// Request parameters
	};
	// var resp;
	
	var body = JSON.stringify({
	  "documents": [
		{
		  "language": "",
		  "id": "1",
		  "text": description
		}
		]
	});
  
	var xhr = new XMLHttpRequest();
	// xhr.withCredentials = true;

	xhr.addEventListener("readystatechange", function () {
	  if (this.readyState === 4) {
		// alert(this.responseText);
		var resp = JSON.parse(this.responseText);
		
		if (type == "language"){
			var languageDetected = resp.documents[0].detectedLanguages[0].name.toString();
			Xrm.Page.getAttribute("new_language").setValue(languageDetected);
		} else if (type == "keyphrase"){
			var keyPhrases = resp.documents[0].keyPhrases.toString();
			Xrm.Page.getAttribute("new_keyphrases").setValue(keyPhrases);
		} else {
			var score  =  resp.documents[0].score.toFixed(2);
			var sentiment  = classifySentimentScore(score);
			Xrm.Page.getAttribute("new_sentimentscore").setValue(score);
			Xrm.Page.getAttribute("new_sentiment").setValue(sentiment);
			setSentimentEmoticon();
		}
	  }
	});
	
	var URL = retrieveConfig("Text Analytics API");
	var subscriptionKey = retrieveConfig("Text Analytics Subscription Key");
		
	if (type == "language"){
		URL = URL + "/languages";		
	} else if (type == "keyphrase"){
		URL = URL + "/keyPhrases";		
	} else {
		URL = URL + "/sentiment";		
	}
	xhr.open("POST", URL);
	xhr.setRequestHeader("Ocp-Apim-Subscription-Key", subscriptionKey);
	xhr.setRequestHeader("Content-Type", "application/json");
	xhr.setRequestHeader("Accept", "application/json");
	// xhr.setRequestHeader("Cache-Control", "no-cache");
	// xhr.setRequestHeader("Postman-Token", "04f0b8ba-474c-ff2e-c3ab-707c6d309575");

	xhr.send(body);
	
}
	
function analyzeCaseDescriptionOnSave(){
	debugger;
	var Description =  Xrm.Page.getAttribute("description").getValue();
	// Xrm.Page.data.entity.addOnSave(callTextAnalyticsAPI(Description));
	callTextAnalyticsAPI("language", Description);
	callTextAnalyticsAPI("keyphrase", Description);
	callTextAnalyticsAPI("sentiment", Description);	
}

function classifySentimentScore(score){
	var sentiment = "";
	
	if (score >= 0.8){
		sentiment = "Very High";
		Xrm.Page.getAttribute("new_satisfaction").setValue(100000000);
	} else if (score >= 0.6){
		sentiment = "High";
		Xrm.Page.getAttribute("new_satisfaction").setValue(100000001);
	} else if (score >= 0.4){
		sentiment = "Neutral";
		Xrm.Page.getAttribute("new_satisfaction").setValue(100000002);
	} else if (score >= 0.2){
		sentiment = "Low";
		Xrm.Page.getAttribute("new_satisfaction").setValue(100000003);
	} else {
		sentiment = "Very Low";
		Xrm.Page.getAttribute("new_satisfaction").setValue(100000004);
	}
	return sentiment;
}

function setSentimentEmoticon(){
	var sentiment = Xrm.Page.getAttribute("new_sentiment").getValue();
	var url = Xrm.Page.context.getClientUrl();
	var iFrame = Xrm.Page.getControl('IFRAME_Sentiment');

	if (sentiment != null){
		if (sentiment == "Very High" || sentiment == "High"){
			iFrame.setSrc(url + "/%7B636541873040000127%7D/WebResources/new_happy.png");
		} else if (sentiment == "Low" || sentiment == "Very Low"){			
			iFrame.setSrc(url + "/%7B636541873040000127%7D/WebResources/new_sad.png");
		} else {
			iFrame.setSrc(url + "/%7B636541873040000127%7D/WebResources/new_neutral.png");
		} 
	}
}

function retrieveConfig(configName){
	var new_value;
	var req = new XMLHttpRequest();
	req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/new_configurationsettings?$filter=new_name eq '" + configName + "'", false);
	req.setRequestHeader("OData-MaxVersion", "4.0");
	req.setRequestHeader("OData-Version", "4.0");
	req.setRequestHeader("Accept", "application/json");
	req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
	req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
	req.onreadystatechange = function() {
		if (this.readyState === 4) {
			req.onreadystatechange = null;
			if (this.status === 200) {
				var results = JSON.parse(this.response);
				for (var i = 0; i < results.value.length; i++) {
					new_value = results.value[i]["new_value"];					
				}
			} else {
				Xrm.Utility.alertDialog(this.statusText);
			}
		}
	};
	req.send();
	return new_value;	
}

function retrieveSurveyResponseComments(){
	var sentiment = Xrm.Page.getAttribute("new_sentiment").getValue();
	//if (sentiment == null){
		var SurveyResponseId = Xrm.Page.data.entity.getId().replace('{', '').replace('}', '');
		var url = Xrm.Page.context.getClientUrl() + "/api/data/v8.2/msdyn_questionresponses?$select=msdyn_name&$filter=msdyn_valueasstring ne null and  _msdyn_surveyresponseid_value eq " + SurveyResponseId;
		var req = new XMLHttpRequest();
		req.open("GET", url, false);
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
		req.onreadystatechange = function() {
			if (this.readyState === 4) {
				req.onreadystatechange = null;
				if (this.status === 200) {
					var results = JSON.parse(this.response);
					for (var i = 0; i < results.value.length; i++) {
						var msdyn_name = results.value[i]["msdyn_name"];
						callTextAnalyticsAPI("language", msdyn_name);
						callTextAnalyticsAPI("keyphrase", msdyn_name);
						callTextAnalyticsAPI("sentiment", msdyn_name);
						Xrm.Page.data.entity.save();					
					}
				} else {
					Xrm.Utility.alertDialog(this.statusText);
				}
			}
		};
		req.send();
	// } else {
		// setSentimentEmoticon();
	// }
}