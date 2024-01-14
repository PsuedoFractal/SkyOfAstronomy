// Import necessary libraries/modules
var fs = require("fs"), // File System module for working with files
    path = require("path"), // Path module for handling file paths
    _ = require("underscore"), // Underscore library for utility functions
    mustache = require("mustache"), // Mustache library for templating
    mammoth = require("mammoth"), // Mammoth library for converting docx to HTML
    cheerio = require("cheerio"); // Cheerio library for parsing HTML

const mjml2html = require('mjml'); // MJML library for converting MJML to HTML

// Entry point of the program
main();

// Function to get the most recent file name from a directory
function getMostRecentFileName(dir) {
    var files = fs.readdirSync(dir); // Read the list of files in the directory
    return _.max(files, function (f) { // Find the file with the latest modification time
        var fullpath = path.join(dir, f);
        return fs.statSync(fullpath).mtime;
    });
}

// Main function
async function main() {
    // Get the most recent file name from the input articles directory
    var flnm = getMostRecentFileName("./input-articles/");
    console.log(flnm); // Display the most recent file name    

    const options = {
        styleMap: [
            "p[style-name='Subtitle'] => h6:fresh"
        ]
    }


    // Convert the docx file to HTML using Mammoth
    const result = await mammoth.convertToHtml({ path: "./input-articles/" + flnm }, options);

    var html = result.value; // Extract the HTML content from the result
    var tempHTML = cheerio.load(html); // Load the HTML content into Cheerio

    html = tempHTML.html(); // Get the HTML content from Cheerio

    const warnings = result.messages; // Get any conversion warnings
    if (warnings.length > 0) {
        console.log("Warnings while converting docx to html:");
        for (var warning of warnings) {
            console.log(warning.message); // Display conversion warnings
        }
    }

    const suffix = await fs.readFileSync('./templates/suffix.html', 'utf8');

    tempHTML = cheerio.load(html);
    tempHTML("img").attr("style", "max-width: 80%; height: auto; margin-left: auto; margin-right: auto; display: block;");
    tempHTML("div").attr("style", "background:linear-gradient(rgb(0,0,0) 0%,rgb(10,26,57) 50%,rgb(18,46,102) 85%,rgb(24,62,134) 100%);");
    tempHTML("h1").attr("style", "font-family: Lato; font-size: 30px; font-weight: 600; line-height: 1.6; text-align: center; color: #ffffff; padding: 0.5em;");
    tempHTML("h2").attr("style", "font-family: Lato; color: #e8b343; font-size: 24px; text-decoration: underline; line-height:1.6;");
    tempHTML("h3").attr("style", "font-family: Lato; color: #ffffff; font-size: 24px;");
    tempHTML("h4").attr("style", "font-family: Lato; color: #ffffff; font-size: 20px;");
    tempHTML("h5").attr("style", "font-family: Lato; color: #ffffff; font-size: 18px;");
    tempHTML("a").attr("style", "font-family: Lato; color: #ff683a; font-size: 16px; line-height: 1.6; text-decoration: underline;");
    tempHTML("p").attr("style", "font-family: Lato; color: #ffffff; font-size: 16px; line-height: 1.6;");
    tempHTML("li").attr("style", "font-family: Lato; color: #ffffff; line-height: 2.5; font-size: 16px;");
    tempHTML("h6").attr("style", "font-family: Lato; border-bottom: 1px solid #808080; font-size: 16px; line-height: 1.5; margin-bottom: 0.5em; padding-bottom: 0.5em; text-align: center; width: 100%;");

    html = tempHTML.html(); // Get the HTML content from Cheerio
    const completeHtml = `
    <body>
    <div style = "width: 100%; padding-left: 20%; padding-right: 20%;background:linear-gradient(rgb(0,0,0) 0%,rgb(10,26,57) 50%,rgb(18,46,102) 85%,rgb(24,62,134) 100%);">
    ${html}
    ${suffix}
    </div>
    </body>
    `;
    
    // Write the HTML content to an output file
    await fs.writeFileSync('./output-articles/mammoth/' + flnm + '.html', completeHtml, 'utf8');

    const content = { content: completeHtml };
    const template = fs.readFileSync('./templates/empty.mjml', 'utf8');

    // Render the MJML template with the HTML content
    const outputMjml = mustache.render(template, content);
    // Write the rendered MJML to an output file
    await fs.writeFileSync('./output-articles/mjml/' + flnm + '.mjml', outputMjml, 'utf8');

    // Convert the MJML to HTML using MJML library
    const output = mjml2html(outputMjml);
    // Write the HTML output to an output file
    await fs.writeFileSync('./output-articles/html/' + flnm + '.html', output.html, 'utf8');
    if (output.errors.length > 0) {
        console.log("Errors while converting mjml to html:");
        for (var error of output.errors) {
            console.log(error.message); // Display conversion errors
        }
    }
}