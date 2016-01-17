var expect = require('chai').expect;
require('mocha-sinon');
var outputData = require('./xwbp');

var xlsxFile = './temp/test.xlsx';

var temp = {
  "Sheet1": [{
    "Username": "jonsmith",
    "FirstName": "John",
    "LastName": "Smith"
  }]
};

temp = JSON.stringify(temp, 2, 2);

describe('xwbp', function() {

  beforeEach(function() {
    this.sinon.stub(console, 'log');
  });

  it('should log -> \n\n' + temp, function() {
    outputData.toJson(xlsxFile);
    expect(console.log.calledOnce).to.be.true;
    expect(console.log.calledWith(temp)).to.be.true;
  });

});

