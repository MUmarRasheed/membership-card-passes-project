const axios = require("axios");
const xlsx = require("xlsx");
const fs = require("fs");
require("dotenv").config();

const API_KEY = process.env.API_KEY;
const MODEL_ID = process.env.MODEL_ID;
const EXCEL_FILE = "members.xlsx";

// Function to read Excel file
function readExcel() {
  const workbook = xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet, { defval: "" }); // Keep empty cells
  return { workbook, sheet, data };
}

// Function to save updated data back to Excel file
function writeExcel(updatedData, workbook, sheet) {
  const updatedSheet = xlsx.utils.json_to_sheet(updatedData, {
    skipHeader: false,
  });
  workbook.Sheets[workbook.SheetNames[0]] = updatedSheet;
  xlsx.writeFile(workbook, EXCEL_FILE);
}

// Function to convert Excel date serial to a valid Date object
function convertExcelDate(serialDate) {
  const excelBaseDate = new Date(1900, 0, 1); // Excel's base date (January 1, 1900)
  const dateOffset = serialDate - 1; // Excel treats 1900-01-01 as day 1
  return new Date(excelBaseDate.getTime() + dateOffset * 24 * 60 * 60 * 1000);
}

// Function to format a Date object to the required ISO string format
function formatToISOWithTimezone(date) {
  const offset = date.getTimezoneOffset();
  const sign = offset > 0 ? "-" : "+";
  const absOffset = Math.abs(offset);
  const hours = String(Math.floor(absOffset / 60)).padStart(2, "0");
  const minutes = String(absOffset % 60).padStart(2, "0");

  // Format date to YYYY-MM-DDTHH:mm:ss
  const isoString = date.toISOString().replace("Z", `${sign}${hours}:${minutes}`);
  return isoString;
}

// Function to convert DD/MM/YYYY date to a valid Date object if it is in string format
function parseExpirationDate(expirationValue) {
  console.log("🚀 ~ parseExpirationDate ~ typeof expirationValue:", typeof expirationValue);

  if (typeof expirationValue === "number") {
    // Convert Excel serial date to JavaScript Date object
    return convertExcelDate(expirationValue);
  }

  if (typeof expirationValue === "string") {
    const [day, month, year] = expirationValue.split("/"); // Expecting DD/MM/YYYY
    if (day && month && year) {
      // Create a date assuming DD/MM/YYYY format
      return new Date(
        `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`
      );
    }
  }
  return null;
}

// Function to create a pass for a single member
async function createPass(member) {
  console.log("🚀 ~ createPass ~ member:", member);
  console.log(
    `Processing member: ${member.Name}, Expiration_Date: ${member.Expiration_Date}`
  );

  try {
    const expirationDate = parseExpirationDate(member.Expiration_Date);
    if (!expirationDate || isNaN(expirationDate.getTime())) {
      console.error(
        `Invalid expiration date for ${member.Name}: ${member.Expiration_Date}`
      );
      return null;
    }

    const expirationDateISO = formatToISOWithTimezone(expirationDate);
    const payload = {
      expirationDate: String(expirationDateISO),
      fields: [
        { key: "field3", value: member.Name },
        { key: "field6", value: String(member.License_Number) },
        { key: "field10", value: String(member.ID_Number) },
        { key: "field7", value:  String(member.Expiration_Date)},
      ],
      images: [], // Will fill this later
    };

    // Upload image if it exists
    if (member.Photo) { // Assuming PhotoUrl is the property with the image URL
      const imageHex = await uploadImage(member.Photo); // Pass the image URL here
      if (imageHex) {
        payload.images.push({ type: "thumbnail", hex: imageHex }); // Use the hex value in the payload
      } else {
        console.error(`Failed to upload image for ${member.Name}`);
      }
    }

    const response = await axios.post(
      `https://api.pass2u.net/v2/models/${MODEL_ID}/passes`,
      payload,
      {
        headers: {
          "x-api-key": API_KEY,
          Accept: "application/json",
          "Content-Type": "application/json",
        },
      }
    );

    console.log("🚀 ~ createPass ~ response:", response.data);
    const { passId } = response.data;
    const passUrl = `https://www.pass2u.net/d/${passId}`;
    console.log(`Created pass for ${member.Name}: ${passUrl}`);
    return passUrl;

  } catch (error) {
    console.error(`Failed to create pass for ${member.Name}`, error.response.data || error.message);
    return null;
  }
}


const uploadImage = async (imageUrl) => {
  console.log("🚀 ~ uploadImage ~ imageUrl:", imageUrl)
  try {
    // Fetch the image from the URL
    const imageResponse = await axios.get(imageUrl, {
      responseType: 'arraybuffer', // Get the response as binary data
    });

    // Upload the image to Pass2U API
    const uploadResponse = await axios.post('https://api.pass2u.net/v2/images', imageResponse.data, {
      headers: {
        "x-api-key": API_KEY,
        Accept: "application/json",
        "Content-Type": "image/png",
      },
    });

    console.log("Image uploaded successfully:", uploadResponse.data);
    return uploadResponse.data.hex; // Return the hex value
  } catch (error) {
    console.error("Failed to upload image:", error.response.data || error.message);
    return null;
  }
};

// Main function to process all members
async function processMembers() {
  const { workbook, sheet, data: members } = readExcel();

  // Add Pass URL column if it doesn't exist
  if (!members[0].hasOwnProperty("Pass_URL")) {
    members.forEach((member) => (member.Pass_URL = "")); // Initialize empty "Pass_URL" field
  }

  for (const member of members) {
    if (!member.Pass_URL) {
      // Skip if the URL already exists
      const passUrl = await createPass(member);
      if (passUrl) {
        member.Pass_URL = passUrl; // Add pass URL to the member
      }
    }
  }

  // Write updated members data back to Excel
  writeExcel(members, workbook, sheet);
  console.log("Pass URLs saved in Excel file");
}

// Start the process
processMembers();
