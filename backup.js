// Function to generate responses using OpenAI's API
async function generateResponses(message) {
  try {
    // https://api.openai.com/v1/engines/davinci/completions
    const response = await axios.post(
      "https://api.openai.com/v1/chat/completions",
      {
        model: "gpt-3.5-turbo",
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: message },
        ],
        // max_tokens: 200, // Adjust this based on the desired response length
        // n: 2, // Number of responses to generate
        temperature: 0.5, // Adjust the temperature for diversity in responses
      },
      {
        headers: {
          Authorization: `Bearer ${OPENAI_API_KEY}`,
          "Content-Type": "application/json",
        },
      }
    );

    const replies = response.data.choices.map((choice) => {
    //   console.log(choice);
      return choice?.['message']?.['content'];
    });
    return replies;
  } catch (error) {
    console.error("Error:", error.message);
    return [];
  }
}

// Use this function to generate responses for a message
async function processMessage(message) {
  const responses = await generateResponses(message);

  // Print the generated responses
//   console.log("Generated responses:");
  responses.forEach((reply, index) => {
    console.log(`Response ${index + 1}: ${reply}`);
  });

  return responses;
}

// Function to perform processing based on column A and write to column C
async function processExcelData() {
  const range = XLSX.utils.decode_range(worksheet["!ref"]);

  for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
    const cellA = "A" + (rowNum + 1);
    const cellB = "B" + (rowNum + 1);
    const cellC = "C" + (rowNum + 1);

    // Get value from column A
    const valueA = worksheet[cellA] ? worksheet[cellA].v : "";
    console.log("Message:  ", valueA);

    // Perform your processing based on valueA
    // For example, let's say we want to double the value from column A
    const processedValue = await processMessage(valueA); // valueA + "sunil"; // Replace this with your processing logic
    // console.log(processedValue);

    // Write the processed value to column C
    worksheet[cellC] = { v: processedValue };
  }

  // Write back to the Excel file
  XLSX.writeFile(workbook, "output.xlsx");
}

// Call the function to process data
processExcelData();
