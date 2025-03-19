const fs = require('fs');
const { parseString, Builder } = require('xml2js');
const path = require('path');
const dir = process.argv[2];
// Function read  directory
function processSVGFilesInDirectory(directory) {
    fs.readdir(directory, (err, files) => {
        if (err) {
            console.error('Error reading directory:', err);
            return;
        }
        // Filter SVG files
        const svgFiles = files.filter(file => path.extname(file).toLowerCase() === '.svg');
        console.log('Found SVG files:', svgFiles);

        // Process each SVG file
        svgFiles.forEach(svgFile => {
            const filePath = path.join(directory, svgFile);
            console.log(svgFile)
            processSVGFile(filePath);
        });
    });
}

// Function  process each SVG file
function processSVGFile(filePath) {
    // Read the SVG file
    fs.readFile(filePath, 'utf-8', (err, data) => {
        if (err) {
            console.error('Error reading SVG file:', err);
            return;
        }
        // Parse SVG content
        parseString(data, (err, result) => {
            if (err) {
                console.error('Error parsing SVG content:', err);
                return;
            }

            // Find the group element by checking custom props
            const groups = result.svg.g[0].g;
            //console.log('Number of groups:', groups.length);
            for (let i = 0; i < groups.length; i++) {
                const g = groups[i];
                //console.log('Group:', g);
                if (g['v:custProps'] && Array.isArray(g['v:custProps'])) {
                    const custProps = g['v:custProps'][0]['v:cp'];
                    //console.log('Custom props:', custProps);
                    const hasLinkProp = custProps.some(cp => cp.$ && cp.$.hasOwnProperty('v:nameU') && cp.$['v:nameU'].toLowerCase() === 'link');
                    const hasTraceProp = custProps.some(cp => cp.$ && cp.$.hasOwnProperty('v:nameU') && cp.$['v:nameU'].toLowerCase() === 'trace');
                    if (hasLinkProp && hasTraceProp) {
                        // Move custom properties from group to shape
                        //console.log(filePath)
                        //console.log(g.g)
                        if (Array.isArray(g.g)) {
                            const shape = g.g.find(shape => shape.$.id.startsWith('shape') && shape.hasOwnProperty('path'));
                            if (shape) { //console.log(shape)
                                if (!shape.hasOwnProperty('v:custProps')) {
                                    shape['v:custProps'] = [];
                                }
                                if (g['v:custProps']) {
                                    g['v:custProps'].forEach(prop => {
                                        const copiedProp = {'v:cp': prop['v:cp'].map(item => ({...item}))};
                                        shape['v:custProps'].push(copiedProp);
                                    });
                                    // Copy title from group to shape
                                    shape.title = g.title;

                                }
                                //console.log('Custom props copied from group to shape:', shape);
                            } else {
                                console.error('Shape element not found within the group');
                            }
                        }
                    }
                }
            }
            // Convert the updated XML object back to string
            const builder = new Builder();
            const updatedSvg = builder.buildObject(result);

            // Write the updated SVG content back to the file
            const outputFile = filePath; //.replace('.svg', '_updated.svg');
            fs.writeFileSync(outputFile, updatedSvg, 'utf-8', err => {
                if (err) {
                    console.error('Error writing updated SVG file:', err);
                    return;
                }
                console.log('SVG file updated successfully:', outputFile);
            });
        });
    });
}


// Specify the directory containing SVG files
const directory = 'C:\\Users\\SHUTEW\\Desktop\\VToR\\untitled\\'+ dir+'\\svg';

// Process SVG files in the directory
processSVGFilesInDirectory(directory);
