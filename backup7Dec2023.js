const XLSX = require("xlsx");
const axios = require("axios");

const BATCH_SIZE = 1;
const API_TIMEOUT = 10000; // 10 seconds timeout for API response

const OPENAI_API_KEY = "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"; // Replace with your OpenAI API key

const SYSTEM_PROMPT = `I will give you a message received from a user.
Analyse it and provide 2 to 3 potential appropriate responses/replies in less than 15 words in response to the message.
 
Below is an example for you in which there are two things one is the original message received and then the reply to it.

message: Hi Andrea, this is the message I received from Professor Harden regarding the ↵chart:↵↵↵↵"Students download the word doc and either place the x in, or close to, the ↵box. She can download it, mark it with pen, and then scan and upload. She can ↵also copy and paste into a new document and select a box by placing an x in the ↵box. The word doc must accept that as that seems to be what most students do. ↵She can also just restate the↵↵words by typing if she can’t access the box. I just need her to somehow let ↵me know what her answers are for the section she omitted."↵↵
 
reply: Thank you so much for your help. I will try these different things and get it ↵done today. ↵↵ ↵↵Andrea Willis ↵↵`;

const workbook = XLSX.readFile("input.xlsx");
const sheetName = "Sheet1"; // Replace with your sheet name
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
    console.log("Sending request to OpenAI API...");

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
        timeout: API_TIMEOUT,
      }
    );

    const replies = response.data.choices.map((choice) => choice?.['message']?.['content']);
    console.log("Received responses from OpenAI:", replies);

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

    if (valueA.trim() !== "") {
      batchPromises.push(processMessageWithDelay(valueA).then((processedValue) => {
        const concatenatedResponse = processedValue.join('\n\n\n');
        console.log(concatenatedResponse);
        worksheet[cellC] = { v: concatenatedResponse };

        // Write back to the output file after each processed row
        XLSX.writeFile(workbook, "output.xlsx");
      }));
    }
  }

  return Promise.all(batchPromises);
}

async function processExcelData() {
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  const totalRows = range.e.r - range.s.r + 1;

  console.log("Total rows in the Excel file:", totalRows);

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

    // Add a delay between batches to stay within API limits
    const delayBetweenBatches = 3000;
    await new Promise((resolve) => setTimeout(resolve, delayBetweenBatches));
  }

  console.log('All rows processed. Output file ready.');
}

// Call the function to process data
processExcelData();