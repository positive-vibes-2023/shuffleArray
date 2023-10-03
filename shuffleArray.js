/* This programme takes user sample data in Excel file format and then selects a
sample of random samples. The size of the random sample is decided by the user*/
/* Please ensure xlsx package is installed using npm environment */
const XLSX = require('xlsx'); /* package imported */
const fs = require('fs'); /* package imported */

const workBook = XLSX.readFile(
	'/Users/gaurav/Library/CloudStorage/OneDrive-NHS/Clinical governance/Audits/CTPA-Audit-final-data.xlsx'
); /* File imported */
const sheetName = workBook.SheetNames[0];
const sheet =
	workBook.Sheets[
		sheetName
	]; /* The first worksheet containing population data imported */
const populationData =
	XLSX.utils.sheet_to_json(
		sheet
	); /* The worksheet data is converted to an object containing key/value pairs */

/* This function shuffles an array using the Fisher-Yates algorithm. */
function shuffle(array) {
	for (let i = array.length - 1; i > 0; i--) {
		let j = Math.floor(Math.random() * (i + 1));
		[array[i], array[j]] = [array[j], array[i]];
	}
	return array;
}
// Call the shuffle function to shuffle the population data
shuffledPopulationData = shuffle(populationData);

//Now reduce the data size to 100 by selecting 100 random samples
sampleData = shuffledPopulationData.slice(0, 100);
// console.log('This is sample data of randomly selected 100 samples: \n ');
// console.log('Sample size: ' + sampleData.length);
// console.log(sampleData);

//Sort the data by indices
sampleData.sort((a, b) => {
	return a.Index - b.Index;
});
console.log('**********************');
console.log('Sample size: ' + sampleData.length);
console.log(sampleData);

/*Create a new worksheet with new sample data*/

const newWorksheet = XLSX.utils.json_to_sheet(sampleData);
/* Add new worksheet to the our existing excel file as new worksheet */
XLSX.utils.book_append_sheet(workBook, newWorksheet, 'Randomized Sample Data');
/* Write the updated workbook back to the Excel file */

XLSX.writeFile(
	workBook,
	'/Users/gaurav/Library/CloudStorage/OneDrive-NHS/Clinical governance/Audits/CTPA-Audit-final-data.xlsx'
);
console.log('all done!');
