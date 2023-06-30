const express = require('express');
const fileUpload = require('express-fileupload');
const { exec } = require('node:child_process');
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const bodyParser = require('body-parser');
const fsP = require('fs').promises;

const app = express();

app.use(express.static('public'));
app.use(fileUpload());
// app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.urlencoded({ extended: true }));

const directoryTemplates = path.join(__dirname, `${process.env.PYTHON_TEMPLATES_PATH}`);
const directorySettings = path.join(__dirname, `${process.env.SAVED_SETTINGS_FILES_PATH}`);
const directoryOriginalFiles = path.join(__dirname, `${process.env.SAVED_FILES_PATH}`);
const directoryModifyFiles = path.join(__dirname, `${process.env.PYTHON_SAVED_FILES_PATH}`);
const directoryErrors = path.join(__dirname, `${process.env.SAVED_ERRPR_PATH}`);

const checkAuthorization = (req, res, next) => {
  // Получение токена из запроса
  const authorizationHeader = req.headers.authorization;
  let token = null;
  if (authorizationHeader) {
    token = authorizationHeader.replace(/^Bearer\s+/, "");
  }

  // Проверка наличия токена
  if (!token) {
    return res.sendFile(__dirname + '/index.html');
  }

  // Проверка валидности токена (вы можете использовать ваш метод проверки токена здесь)
  if (token !== process.env.TOKEN) {
    return res.sendFile(__dirname + '/index.html');
  }
  
  // Продолжение выполнения следующего middleware или основного запроса
  next();
};

app.use(checkAuthorization);

function generateId() {
  const prefix = "id-";
  const randomNumber = Math.random().toString(36).substring(2);
  return prefix + randomNumber;
}

app.get('/templates', (req, res) => {
  fs.readdir(directoryTemplates, function (err, templates) {
    if (err) {
      res.status(200).json([]);
      return console.log('Unable to scan directory: ' + err);
    }
    res.status(200).json(templates);
  });
});

app.get('/settings', (req, res) => {
  fs.readdir(directorySettings, function (err, settings) {
    if (err) {
      res.status(200).json([]);
      return console.log('Unable to scan directory: ' + err);
    }

    const result = settings.filter(file => {
      if (file.endsWith('.json')) {
        return true;
      } else {
        return false;
      }
    });
    res.status(200).json(result);
  });
});

app.post('/setting', async (req, res) => {
  const file = req.body.name;
  if (file.endsWith('.json')) {
    try {
      const data = await fsP.readFile(directorySettings + '/' + file, 'utf8');
      res.status(200).json({ result: JSON.parse(data) });
    } catch (err) {
      console.log('Could not read file:', file, err);
      res.status(200).json({ result: null });
    }
  }
});

// app.get('/setting', (req, res) => {
//   fsP.readdir(directorySettings)
//     .then(files => {
//       const readFilePromises = files.map(async file => {
//         if (file.endsWith('.json')) {
//           try {
//             const data = await fsP.readFile(directorySettings + '/' + file, 'utf8');
//             return JSON.parse(data);
//           } catch (err) {
//             console.log('Could not read file:', file, err);
//             return null;
//           }
//         }
//         return null;
//       });
//       Promise.all(readFilePromises)
//       .then(dataArray => {
//         const filteredDataArray = dataArray.filter(data => data !== null);
//         res.status(200).json(filteredDataArray);
//       })
//       .catch(err => {
//         console.log('Could not read files:', err);
//         res.status(500).send('Server error');
//       });
//     })
//     .catch(err => {
//       res.status(200).json([]);
//       console.log('Could not read the directory:', err);
//     });
// });

app.post('/new_setting', (req, res) => {
  res.status(200).json({ result: true });
  fs.writeFileSync(directorySettings + `/${req.body.title}` + '.json', JSON.stringify(req.body))
  // const fs = require('fs');
  // const data = fs.readFileSync('file.json', 'utf8');
  // const obj = JSON.parse(data);
});

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});

app.post('/python', (req, res) => {
  if (!req.files || !req.files.file) {
    return res.status(400).send('No file uploaded');
  }

  const template = req.body.template;

  const { name, data } = req.files.file;

  const newName = generateId() + '.xlsx';

  try {
    // Сохраняем загруженный файл на диск
    fs.writeFileSync(directoryOriginalFiles + `/${newName}`, data);
  } catch (err) {
    console.error(err);
  }

  if (!template) {
    template = "default_template";
  }

  const script = `python3 ${directoryTemplates}/${template} ` + `${newName}`;
  // Выполняем скрипт Python с передачей имени файла в качестве аргумента
  exec(script, (error, stdout, stderr) => {
    if (error) {
      console.error(`Error: ${error.message}`);
      res.status(200).json({ result: error.message });
    } else if (stderr) {
      console.error(`Error: ${stderr}`);
      res.status(200).json({ result: stderr });
    } else {
      if (stdout === 'False\n') {
        const result = 'Входные данные не соответствуют скрипту';
        res.status(200).json({ result: result });
        res.on('finish', () => {
          try {
            // Удаляем файлы с диска
            fs.unlinkSync(directoryOriginalFiles + `/${newName}`);
            console.log('File deleted');
          } catch (err) {
            console.error(err);
          }
        });
      } else if (stdout === 'True\n') {
        const filePath = directoryModifyFiles + `/${newName}`;
        const filePathHtml = filePath.split('.')[0] + '.html'; 
        fs.readFile(filePathHtml, 'utf8', (err, html) => {
          if (err) {
            res.status(500).send('Error reading file');
          } else {
            res.status(200).json({ result: html, filename: newName });
            res.on('finish', () => {
              try {
                // Удаляем файлы с диска
                fs.unlinkSync(directoryOriginalFiles + `/${newName}`);
                fs.unlinkSync(filePathHtml);
                console.log('File deleted');
              } catch (err) {
                console.error(err);
              }
            });
          }
        });
      } else if (stdout === 'unknowns_division\n') {
        const filename = 'unknowns_division.json';
        const filePath = path.join(directoryErrors, filename);
        fs.readFile(filePath, 'utf8', (err, data) => {
          if (err) {
            res.status(500).send('Error reading file');
          } else {
            data = JSON.parse(data);
            let result = 'Нет дивизиона у: ';
            for (let key in data.error) {
              result = result + data.error[key]
            }
            res.status(200).json({ result: result });
            res.on('finish', () => {
              try {
                // Удаляем файлы с диска
                fs.unlinkSync(directoryOriginalFiles + `/${newName}`);
                fs.unlinkSync(filePath);
                console.log('File deleted');
              } catch (err) {
                console.error(err);
              }
            });
          }
        });
      } else {
        console.log(stdout)
        res.status(200).json({ result: stdout });
      }

    }
    // res.send(output);
  });
});

app.post('/save_python', (req, res) => {
  if (!req.files || !req.files.file) {
    return res.status(400).send('No file uploaded');
  }

  const { name, data } = req.files.file;

  try {
    // Сохраняем загруженный файл на диск
    fs.writeFileSync(directoryTemplates + `/${name}`, data);
  } catch (err) {
    console.error(err);
  }
});

app.get('/download', function(req, res){
  const filename = req.query.filename;
  const file = path.join(directoryModifyFiles, filename);
  // проверка наличия файла
  fs.access(file, fs.constants.F_OK, (err) => {
    if (err) {
      console.error(err);
      return res.status(404).send('File not found');
    }
    // загрузка файла
    res.download(file, filename, (err) => {
      if (err) {
        console.error(err);
        return res.status(500).send('Failed to download file');
      }
      // удаление файла
      fs.unlink(file, (err) => {
        if (err) {
          console.error(err);
        }
        console.log(`File ${filename} deleted`);
      });
    });
  });
});

app.get('/download_template', function(req, res){
  const filename = req.query.filename;
  const file = path.join(directoryTemplates, filename);
  // проверка наличия файла
  fs.access(file, fs.constants.F_OK, (err) => {
    if (err) {
      console.error(err);
      return res.status(404).send('File not found');
    }
    // загрузка файла
    res.download(file, filename, (err) => {
      if (err) {
        console.error(err);
        return res.status(500).send('Failed to download file');
      }
    });
  });
});

app.listen(process.env.PORT, () => console.log('Server listening on port 3000'));

