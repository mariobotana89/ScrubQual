const chai = require('chai');
const expect = chai.expect;

// Import the validation functions
const {
  validateFirstName,
  validateLastName,
  validateEnrolleeID,
  validateContractID,
  validatePBP,
  validateFirstTierEntity,
  validateAuthOrClaimNumber,
  validateDateReceived,
  validateTimeReceived,
  validatePartBDrugRequest,
  validateAORDate,
  validateAORReceiptTime,
  validateRequestDetermination,
  validateRequestProcessing,
  validateTimeframeExtension,
  validateDateOfDetermination,
  validateTimeOfDetermination,
  validateDateOralNotification,
  validateTimeOfOralNotification,
  validateDateWrittenNotification,
  validateTimeWrittenNotification,
  validateWhoMadeRequest,
  validateIssueDescription,
  validateExpeditedRequest,
  validateWasRequestDenied
} = require('../server'); // Adjust the path as needed

describe('Validation Functions', () => {
  it('should validate first name correctly', () => {
    let errors = [];
    validateFirstName('', 2, errors);
    expect(errors).to.have.lengthOf(1);
    expect(errors[0].error).to.equal('Invalid value. Please ensure the Enrollee First Name field is not left blank.');
    
    errors = [];
    validateFirstName('John1', 2, errors);
    expect(errors).to.have.lengthOf(1);
    expect(errors[0].error).to.equal('Invalid value. The Enrollee First Name field cannot contain numeric values. Please enter a valid first name.');
    
    errors = [];
    validateFirstName('John', 2, errors);
    expect(errors).to.have.lengthOf(0);
  });

  // Add similar tests for other validation functions
});
