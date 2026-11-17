const fs = require('fs');
let xml = fs.readFileSync('G:\\1Click Build\\Source\\Excel_customUI.xml', 'utf8');
xml = xml.split('imageMso="ContactCard"').join('imageMso="AddressBook"');
fs.writeFileSync('G:\\1Click Build\\Source\\Excel_customUI.xml', xml, 'utf8');
console.log('Updated Excel_customUI.xml with AddressBook icon');
