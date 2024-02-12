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

	const prompt = `Taking the information from this form as a JSON object, build a retreat plan as a JSON object, or in other words an "itenerary" for this person and their business and return the retreat plan / itenerary as a JSON object including:
  1. "Agenda": A detailed plan of events and scheduling for the event over the ${dRetreat}
  2. "Tags": Relevant Tags 🏷️
  3. "KeyTakeaways": Key takeaways 📝, and any additional notes for the client on their upcoming event
  4. "ActionItems": Action items and/or suggessted recommendations for their upcoming event
  5. "emailResponse": an email for us to send to ${firstName} ${lastName} from ${companyName} at ${companyEmail} telling them about the Agenda and Itenerary that we have generated for them, including all related ${values} and any other relevant information we can share with them at this time. Remember this must be a JSON object`;

	// Analyze the feedback if present
	if (values) {
		const analysis = callGPTToAnalyzeFeedback(prompt);

		// Send the analysis to Slack with formatting and emojis
		sendToSlack(analysis.toString(), email, companyName);

		// Update the Google Sheet with tags and categories
		updateSheetWithTag(range, analysis);
	}
}

// Function to call GPT4 to analyze the feedback
function callGPTToAnalyzeFeedback(prompt) {
	const url = 'https://api.openai.com/v1/chat/completions';
	const payload = {
		model: 'gpt-4-0613',
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
// Function to send the analysis to Slack
function sendToSlack(analysis, email, companyName) {
	const payload = {
		text: `🌟 *New Event Agenda* 🌟
    
👤 *Form Entry From*: ${email}

🥳 *New Event for ${companyName}*: 
- *Agenda*: ${analysis}`,
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
	const sheet = range.getSheet();
	const row = range.getRow();

	// Update the Google Sheet in new columns
	sheet.getRange(`T${row}`).setValue(analysis);
}
