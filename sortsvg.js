const fs = require('fs');
const xml2js = require('xml2js');
const path = require('path')

// function rearrangeSvg(svgContent) {
//     //let file_path = path.join('C:\\Users\\SHUTEW\\Downloads\\process_adhoc\\output',fname+'.svg');
//     // Read the SVG file
//     //const svgContent = fs.readFileSync(file_path, 'utf8');
//
//     // Parse the SVG content
//     xml2js.parseString(svgContent, { explicitArray: false }, (err, result) => {
//         if (err) {
//             console.error('Error parsing SVG:', err);
//             return;
//         }
//         const svg = result.svg;
//         // Extract <g> elements with at least 3 characters in their id attribute
//         const matchingGroups = svg.g.filter(group => group.$.id && group.$.id.length >= 3);
//
//         // Sort the matching <g> elements based on the last 3 characters of their id attribute
//         const sortedMatchingGroups = matchingGroups.sort((a, b) => {
//             const aId = a.$.id.slice(-3);
//             const bId = b.$.id.slice(-3);
//             return aId.localeCompare(bId);
//         });
//
//         // Remove matching <g> elements from the SVG
//         svg.g = svg.g.filter(group => group.$.id.length < 3);
//
//         // Append the sorted <g> elements back to the SVG
//         svg.g.push(...sortedMatchingGroups);
//         // Find the index of the element with id="background"
//         const backgroundIndex = svg.g.findIndex(group => group.$.id === 'background');
//
//         // Move the element with id="background" to the top
//         if (backgroundIndex !== -1) {
//             const backgroundElement = svg.g.splice(backgroundIndex, 1)[0];
//             svg.g.unshift(backgroundElement);
//         }
//         // Convert the modified object back to XML
//         const builder = new xml2js.Builder();
//         const modifiedSvgContent = builder.buildObject(result);
//
//         // Write the updated SVG content back to the file
//         //fs.writeFileSync(file_path, modifiedSvgContent, 'utf8');
//         return modifiedSvgContent;
//     });
// }

// Example usage: rearrangeSvg('your_svg_file.svg');

function rearrangeSvg(file_path) {
    // Read the SVG file
    const svgContent = fs.readFileSync(file_path, 'utf8');
    // Parse the SVG content
    xml2js.parseString(svgContent, { explicitArray: false }, (err, result) => {
        if (err) {
            console.error('Error parsing SVG:', err);
            return;
        }

        const svg = result.svg;

        // Extract <g> elements with at least 3 characters in their id attribute
        const matchingGroups = svg.g.filter(group => group.$.id && group.$.id.length >= 3);

        // Sort the matching <g> elements based on the last 3 characters of their id attribute
        const sortedMatchingGroups = matchingGroups.sort((a, b) => {
            const aId = a.$.id.slice(-3);
            const bId = b.$.id.slice(-3);
            return aId.localeCompare(bId);
        });

        // Remove matching <g> elements from the SVG
        svg.g = svg.g.filter(group => group.$.id.length < 3);

        // Append the sorted <g> elements back to the SVG
        svg.g.push(...sortedMatchingGroups);

        // Find the index of the element with id="background"
        const backgroundIndex = svg.g.findIndex(group => group.$.id === 'background');

        // Move the element with id="background" to the top
        if (backgroundIndex !== -1) {
            const backgroundElement = svg.g.splice(backgroundIndex, 1)[0];
            svg.g.unshift(backgroundElement);
        }

        // Remove <v:custProps> tags containing only "&#xD;" anywhere in the file
                svg.g.forEach(group => {
                    if (
                        group['v:custProps'] &&
                        group['v:custProps'].length === 1 &&
                        group['v:custProps'][0] === '&#xD;'
                    ) {
                        delete group['v:custProps'];
                    }
                });

        // Convert the modified object back to XML
        const builder = new xml2js.Builder();
        const modifiedSvgContent = builder.buildObject(result);

        // Return the updated SVG content
        fs.writeFileSync(file_path, modifiedSvgContent, 'utf8');
        return modifiedSvgContent;
    });
}


//rearrangeSvg('ksouthtx.3271');



function removeCustProps(filePath) {
    // Read the file content
    let fileContent;

    try {
        fileContent = fs.readFileSync(filePath, 'utf8');
    } catch (err) {
        console.error('Error reading file:', err);
        return;
    }

    // Replace <v:custProps>&#xD;</v:custProps> with an empty string
    const modifiedContent = fileContent.replace(/<v:custProps><\/v:custProps>/g, '');

    // Write the modified content back to the file
    try {
        fs.writeFileSync(filePath, modifiedContent, 'utf8');
        console.log('File updated successfully.');
    } catch (err) {
        console.error('Error writing to file:', err);
    }
}

// Specify the file path
//const filePath = 'C:\\Users\\SHUTEW\\Downloads\\process_adhoc\\output\\knorthtx.3271.svg';
// Call the function
//removeCustProps(filePath);


 //updateSvg('3271','3271_KU_SOUTH_TX');

module.exports={rearrangeSvg,removeCustProps};