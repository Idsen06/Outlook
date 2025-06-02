const axios = require('axios');

const API_KEY = "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";

const sendPrompt = async (prompt, retryCount = 0) => {
  try {
    const messages = [
      { role: 'system', content: 'You are a phishing detection machine and your job is to analyse the provided email for phishing. You should answer what the likelihood of the email being a phishing email is in percentage, the higher, the more likely. An answer if the email is phishing, true if it is, false if it is not. Finally provide an explanation of your decision. Provide your answer in the following JSON format, which is extremely important is being followed: Answer:{"likelihood_percentage": integer,"is_phishing": Boolean,"explanation": String}' },
      { role: 'user', content: prompt }
    ];


    
    //for debugging
    console.log('Messages:', messages);

    const response = await axios.post(
      "https://openrouter.ai/api/v1/chat/completions",
      {
        model: "openai/gpt-4.1",
        messages: messages,
        //max_tokens: 150
      },
      {
        headers: {
          'Authorization': `Bearer ${API_KEY}`,
          'Content-Type': 'application/json'
        }
      }
    );

    //for debugging
    console.log('API response:', response.data);
    console.log('API response:', response.messages);

    if (response.data.choices == null || response.data.choices.length < 1) {
      return { error: 'no response from ChatGPT' };
    }

    const messageContent = response.data.choices[0].message?.content;
    if (!messageContent) {
      return { error: 'no content in response from ChatGPT' };
    }

    //vlidate response format
    if (!isValidResponseFormat(messageContent)) {
      console.warn('Invalid response format, re-prompting...');
      if (retryCount < 4) {
        return await sendPrompt(prompt, retryCount + 1); //re-prompt if the format is invalid
      } else {
        return { error: 'Invalid response format after 4 retries' };
      }
    }

    //parse response content into json object
    const parsedResponse = JSON.parse(messageContent.replace('Answer:', '').trim());
    return parsedResponse;
  } catch (error) {
    console.error('Error sending prompt:', error);
    throw error;
  }
};

//function to validate response format
const isValidResponseFormat = (response) => {
  try {
    const parsedResponse = JSON.parse(response.replace('Answer:', '').trim());
    return (
      typeof parsedResponse.likelihood_percentage === 'number' &&
      typeof parsedResponse.is_phishing === 'boolean' &&
      typeof parsedResponse.explanation === 'string'
    );
  } catch (error) {
    console.error('Error validating response format:', error);
    return false;
  }
};

module.exports = { sendPrompt };