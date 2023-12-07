const XLSX = require("xlsx");
const axios = require("axios");

const BATCH_SIZE = 1;
// const REQUESTS_PER_SECOND = 1 // Adjusted rate to respect the RPD limit
const REQUESTS_PER_SECOND = 3500 / (60 * 24 * 60); // Adjusted rate to respect the RPD limit
const API_TIMEOUT = 10000; // 10 seconds timeout for API response


const SYSTEM_PROMPT = `I will give you a message received from a user.
Analyse it and provide 2 to 3 potential appropriate responses/replies in less than 15 words in response to the message.
 
Below is an example for you in which there are two things one is the original message received and then the reply to it.

message: Hi Andrea, this is the message I received from Professor Harden regarding the ↵chart:↵↵↵↵"Students download the word doc and either place the x in, or close to, the ↵box. She can download it, mark it with pen, and then scan and upload. She can ↵also copy and paste into a new document and select a box by placing an x in the ↵box. The word doc must accept that as that seems to be what most students do. ↵She can also just restate the↵↵words by typing if she can’t access the box. I just need her to somehow let ↵me know what her answers are for the section she omitted."↵↵
 
reply: Thank you so much for your help. I will try these different things and get it ↵done today. ↵↵ ↵↵Andrea Willis ↵↵`;

const workbook = XLSX.readFile("input.xlsx");
const sheetName = "sheet1"; // Replace with your sheet name
const worksheet = workbook.Sheets[sheetName];

// Simulated function to mimic API behavior
async function mockGenerateResponses(message) {
  // Simulate some processing time
  await new Promise(resolve => setTimeout(resolve, 1000));

  // Simulate generating responses
  const responses = [
    "This is a mock response for: " + message,
    "Another mock response for: " + message
  ];

  return responses;
}

// Function to generate responses using OpenAI's API
async function generateResponses(message) {
  try {
    console.log("request 1");
    const response = await axios.post(
      "https://api.openai.com/v1/chat/completions",
      {
        model: "gpt-3.5-turbo",
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: message },
        ],
        temperature: 0.5,
      },
      {
        headers: {
          Authorization: `Bearer ${OPENAI_API_KEY}`,
          "Content-Type": "application/json",
        },
        timeout: API_TIMEOUT, // Set timeout for API requests
      }
    );

    const replies = response.data.choices.map((choice) => {
      return choice?.['message']?.['content'];
    });
    // console.log(replies);
    return replies;
  } catch (error) {
    console.error("Error:", error.message);
    return [];
  }
}

// Use this function to generate responses for a message
async function processMessageWithDelay(message) {
  return new Promise(async (resolve) => {
    // const responses = await mockGenerateResponses(message);
    const responses = await generateResponses(message);
    console.log(responses);
    resolve(responses);
  });
}

async function processBatchOfRows(startRow, endRow) {
  const batchPromises = [];

  for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
    const cellA = "A" + (rowNum + 1);
    const cellC = "C" + (rowNum + 1);

    const valueA = worksheet[cellA] ? worksheet[cellA].v : "";
    // console.log(valueA);

    batchPromises.push(processMessageWithDelay(valueA).then(processedValue => {
      const concatenatedResponse = processedValue.join('\n\n\n');
      console.log(concatenatedResponse);
      worksheet[cellC] = { v: concatenatedResponse };
    }));
  }

  return Promise.all(batchPromises);
}

async function processExcelData() {
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  const totalRows = range.e.r - range.s.r + 1;
  console.log(totalRows);
  
  let processedRows = 0;

  while (processedRows < totalRows) {
    const remainingRows = totalRows - processedRows;
    const currentBatchSize = Math.min(BATCH_SIZE, remainingRows);

    const startRow = processedRows;
    const endRow = processedRows + currentBatchSize - 1;

    await processBatchOfRows(startRow, endRow);

    processedRows += currentBatchSize;

    const currentDate = new Date().toLocaleString();
    console.log(`[${currentDate}] Processed ${processedRows} out of ${totalRows} rows`);

    const delayBetweenBatches = 1000;
    // const delayBetweenBatches = 1000 / REQUESTS_PER_SECOND;
    await new Promise(resolve => setTimeout(resolve, delayBetweenBatches));
  }

  XLSX.writeFile(workbook, "output.xlsx");
  console.log('All rows processed. Output file ready.');
}

// Call the function to process data
processExcelData();