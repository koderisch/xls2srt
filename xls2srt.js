const yargonaut = require('yargonaut');
const yargs = require('yargs');
const xlsx = require('node-xlsx');
const fs = require('fs');
const chalk = require('chalk');
const figlet = require('figlet');

console.log(chalk.cyan(figlet.textSync('XLS 2 SRT', { horizontalLayout: 'full' })));

yargonaut.style('blue').style('blue', 'number').style('yellow', 'required').helpStyle('green.underline').errorsStyle('red.bold');

// set up options for command line
const argv = yargs
  .option('input', {
    alias: 'i',
    description: 'Path to xls(x) file or folder containing xls(x) files to process',
    type: 'string',
    demand: true,
  })
  .option('fps', {
    description: 'Frames per second of the video clip',
    alias: 'f',
    type: 'number',
    demand: true,
  })
  .alias('help', 'h').argv;

class ConvertXls2srt {
  constructor(inputPath, fps) {
    this.inputPath = inputPath;
    this.fps = fps;
  }
  processInputPath() {
    // check if path provided is pointing to a file or directory
    if (fs.existsSync(this.inputPath)) {
      if (fs.lstatSync(this.inputPath).isFile()) {
        //console.log("file");
        // if it is a file, process a single file
        this.convertToSrt(this.inputPath);
      } else if (fs.lstatSync(this.inputPath).isDirectory()) {
        //console.log("dir");
        // if it is a folder, process all xls or xlsx files inside the folder
        let files = fs.readdirSync(this.inputPath).map((fileName) => {
          return fileName;
        });
        files.sort();
        files.forEach((file) => {
          if (file.endsWith('.xlsx') || file.endsWith('.xls')) {
            this.convertToSrt(this.inputPath + '/' + file);
          }
        });
      } else {
        console.log('input path supplied is not a file or directory');
        process.exit(0);
      }
    } else {
      console.log("input path doesn't exist");
      process.exit(0);
    }
  }

  convertToSrt(filePathIn) {
    let srtData = '';
    const filePathOut = filePathIn.split('.xls')[0] + '.srt';

    const workSheetsFromFile = xlsx.parse(filePathIn);
    //console.log(workSheetsFromFile[0]["data"]);
    let rows = workSheetsFromFile[0]['data'];

    for (const row in rows) {
      if (row > 0) {
        let subNo = rows[row][0];
        let subIN = this.convertToMillisecondsTC(rows[row][1]);
        let subOUT = this.convertToMillisecondsTC(rows[row][2]);
        var subTxt = rows[row][4].replace(/\s*$/, '');
        srtData += subNo + '\n' + subIN + ' --> ' + subOUT + '\n' + subTxt + '\n\n';
        //console.log(rows[row][0]);
      }
    }
    //console.log(srtData);

    fs.writeFile(filePathOut, srtData, function (err) {
      if (err) {
        return console.log(err);
      }
      console.log('created ' + filePathOut);
    });
  }

  convertToMillisecondsTC(tc) {
    let frames = tc.substr(9, 2);
    let milliseconds = Math.floor((1000 / this.fps) * frames);
    let formattedMs = ('00' + milliseconds).slice(-3);
    let msTC = tc.substr(0, 9) + formattedMs;
    //console.log(msTC);
    return msTC;
  }
}

let myConvertXls2srt = new ConvertXls2srt(argv.input, argv.fps);
myConvertXls2srt.processInputPath();
