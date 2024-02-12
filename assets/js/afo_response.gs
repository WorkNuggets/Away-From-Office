// Access OPENAI_API_KEY and SLACK_WEBHOOK_URL from the script properties
const OPENAI_API_KEY =
	PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
const SLACK_WEBHOOK_URL =
	PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');

// Function to be triggered when the form is submitted
function onFormSubmit(e) {
	// Get the submitted data
	const { values, range } = e;
	const [
		timestamp,
		email,
		firstName,
		lastName,
		companyEmail,
		companyName,
		retreatStartDate,
		nAttendees,
		dRetreat,
		groupDynamics,
		groupPreferences,
		locationPreferences,
		destinationPreferences,
		roomPreferences,
		hotelPreferences,
		retreatVibe,
		roomReservationPreferences,
		budget,
		vision,
	] = values;

	const prompt = `Taking the information from this form, build a retreat plan in the form of a structured JSON object. The plan should include the following elements with the exact structure:
{
  "AFO": {
    "Agenda": [{}],
    "Tags": [],
    "KeyTakeaways": [],
    "ActionItems": [],
    "emailResponse": {
      "Subject": "",
      "Body": ""
    }
  }
}
Fill in the details for a ${dRetreat}-day retreat for ${firstName} ${lastName} from ${companyName}, including all relevant information that can be shared with them at this time. Ensure the plan is comprehensive and tailored to their business needs.

Here is some additional information about the company to help you make decisions on how to build this out:
- Company Name: ${companyName}
- Retreat Start Date: ${retreatStartDate}
- Number of Attendees: ${nAttendees}
- Number of Days for the Retreat: ${dRetreat}
- Dynamics of the group: ${groupDynamics}
- Things the group likes and does not like: ${groupPreferences}
- Preferred Location Preferences: ${locationPreferences}
- Specific Destinations they have in mind: ${destinationPreferences}
- Room Prefrerences: ${roomPreferences}
- Hotel Preferences: ${hotelPreferences}
- The feel for this retreat should be ${retreatVibe}
- Room Reservation for team meeting preferences: ${roomReservationPreferences}
- The budget for this retreat is in this ballpark of: ${budget}
- The vision for this retreat is ${vision}`;

	// Analyze the feedback if present
	if (values) {
		const analysis = callGPTToAnalyzeFeedback(prompt);
		const formattedAnalysis = JSON.parse(analysis);

		// Send the analysis to Slack with formatting and emojis
		sendToSlack(formattedAnalysis, companyEmail, companyName);

		// Update the Google Sheet with tags and categories
		updateSheetWithTag(range, formattedAnalysis);
	}
}

// Function to call GPT4 to analyze the feedback
function callGPTToAnalyzeFeedback(prompt) {
	const url = 'https://api.openai.com/v1/chat/completions';
	const payload = {
		model: 'gpt-4-1106-preview',
		messages: [
			{
				role: 'system',
				content: "You are a helpful assistant helping to provide agenda and itineraries on behalf of an events company that helps startups and small companies to plan their company offsites. Results are to be passed to Slack via a Slack message which should be detailed and focus on the itinerary and agenda for the client company's upcoming offsite event. In addition to being posted within our company Slack channel, this information is also going to be posted online for the clients to view within our login portal on our site as well.",
			},
			{
				role: 'user',
				content: prompt,
			},
		],
		response_format: { type: 'json_object' },
	};

	const params = {
		method: 'post',
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${OPENAI_API_KEY}`,
		},
		payload: JSON.stringify(payload),
	};

	try {
		const response = UrlFetchApp.fetch(url, params);
		const jsonResponse = JSON.parse(response.getContentText());
		return jsonResponse.choices[0].message.content.trim();
	} catch (error) {
		return `Error: ${error.message}`;
	}
}

// Function to send the analysis to Slack
function sendToSlack(analysis, email, companyName) {
	const { Agenda, Tags, KeyTakeaways, ActionItems, emailResponse } =
		analysis.AFO;
	const payload = {
		text: `🌟 *New Event Agenda* 🌟
    
👤 *Form Entry From*: ${email}

🥳 *New Event for Company: ${companyName}*: 
- *Agenda* 📅: \`${JSON.stringify(Agenda)}\`,

- *Tags* 🏷️: ${Tags.join(', ')},

- *Key Takeaways* 📝: ${KeyTakeaways.join(', ')},

- *Action Items* 🔥: ${ActionItems.join(', ')},

- *Template Email* 📤: ${emailResponse.Body},

- *Email Subject* 📚: ${emailResponse.Subject}`,
	};

	const params = {
		method: 'post',
		payload: JSON.stringify(payload),
	};

	try {
		UrlFetchApp.fetch(SLACK_WEBHOOK_URL, params);
	} catch (error) {
		console.error(`Error sending to Slack: ${error.message}`);
	}
}

// Function to update the Google Sheet with tags and categories
function updateSheetWithTag(range, analysis) {
	const { Agenda, Tags, KeyTakeaways, ActionItems, emailResponse } =
		analysis.AFO;
	const sheet = range.getSheet();
	const row = range.getRow();

	// Update the Google Sheet in new columns
	sheet.getRange(`T${row}`).setValue(JSON.stringify(Agenda));
	sheet.getRange(`U${row}`).setValue(Tags.join(', '));
	sheet.getRange(`V${row}`).setValue(KeyTakeaways.join(', '));
	sheet.getRange(`W${row}`).setValue(ActionItems.join(', '));
	sheet.getRange(`X${row}`).setValue(emailResponse.Subject);
	sheet.getRange(`Y${row}`).setValue(emailResponse.Body);
}
