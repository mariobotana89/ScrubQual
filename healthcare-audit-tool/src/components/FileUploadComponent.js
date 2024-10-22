import React from 'react';
import * as XLSX from 'xlsx'; // Import the xlsx library

function FileUploadComponent({ onFileUpload }) {
  // Validation function
  const validateData = (data) => {
    const errors = []; // To store error messages

    data.forEach((row, index) => {
      if (index === 0) return; // Skip header row

      // Enrollee First Name: Column A (index 0), max 50 chars
      if (row[0] && row[0].length > 50) {
        errors.push(`Row ${index + 1}: Enrollee First Name exceeds 50 characters.`);
      }

      // Enrollee Last Name: Column B (index 1), max 50 chars
      if (row[1] && row[1].length > 50) {
        errors.push(`Row ${index + 1}: Enrollee Last Name exceeds 50 characters.`);
      }

      // Enrollee ID: Column C (index 2), exactly 11 uppercase alphanumeric, no hyphens
      if (!/^[A-Z0-9]{11}$/.test(row[2])) {
        errors.push(`Row ${index + 1}: Enrollee ID is not 11 uppercase alphanumeric characters.`);
      }

      // Contract ID: Column D (index 3), starts with an uppercase letter followed by 4 numbers
      if (!/^[A-Z][0-9]{4}$/.test(row[3])) {
        errors.push(`Row ${index + 1}: Contract ID does not match format.`);
      }

      // Plan Benefit Package (PBP): Column E (index 4), exactly 3 numeric characters
      if (!/^[0-9]{3}$/.test(row[4])) {
        errors.push(`Row ${index + 1}: PBP is not exactly 3 numeric characters.`);
      }

      // First Tier, Downstream, and Related Entity: Column F (index 5), max 70 chars
      if (row[5] && row[5].length > 70) {
        errors.push(`Row ${index + 1}: First Tier, Downstream, and Related Entity exceeds 70 characters.`);
      }

      // Add checks for other columns as needed
    });

    return errors; // Return the array of error messages
  };

  const handleFileChange = (event) => {
    const file = event.target.files[0]; // Get the file
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Read data as a raw array
        
        // Validate the data
        const errors = validateData(jsonData);
        if (errors.length > 0) {
          console.error("Validation errors:", errors);
          onFileUpload(errors); // Pass errors up if needed
        } else {
          console.log("No validation errors found.");
          onFileUpload(jsonData); // Pass the valid data up if needed
        }
      };
      reader.readAsBinaryString(file);
    }
  };

  return (
    <div>
      <input type="file" accept=".xlsx" onChange={handleFileChange} />
    </div>
  );
}

export default FileUploadComponent;
