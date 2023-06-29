const path = require('path');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const pythonSavedFiles = path.resolve(__dirname, 'python_saved_files');
const pythonTemplates = path.resolve(__dirname, 'python_templates');
const savedErrorsFiles = path.resolve(__dirname, 'saved_errors_files');
const savedFiles = path.resolve(__dirname, 'saved_files');
const savedSettingsFiles = path.resolve(__dirname, 'saved_settings_files');
const utils = path.resolve(__dirname, 'utils');
const indexHtml = path.resolve(__dirname, 'index.html');
const env = path.resolve(__dirname, '.env');

module.exports = {
    entry: './index.js',
    output: {
      filename: 'excel-parser.js',
      path: path.resolve(__dirname, 'dist')
    },
    mode: 'production',
    target: 'node',
    module: {
      rules: [
        {
          test: /\.(html|env)$/,
          use: [
            {
              loader: 'file-loader',
              options: {
                name: '[name].[ext]'
              }
            }
          ]
        },
        {
          test: /python_(saved|templates|settings|errors)_files\//,
          use: [
            {
              loader: 'file-loader',
              options: {
                name: 'python_[folder]/[name].[ext]'
              }
            }
          ]
        }
      ]
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          {
            from: pythonSavedFiles,
            to: 'python_saved_files',
          },
          {
            from: pythonTemplates,
            to: 'python_templates',
          },
          {
            from: savedErrorsFiles,
            to: 'saved_errors_files',
          },
          {
            from: savedFiles,
            to: 'saved_files',
          },
          {
            from: savedSettingsFiles,
            to: 'saved_settings_files',
          },
          {
            from: utils,
            to: 'utils',
          },
          {
            from: indexHtml,
            to: '.'
          },
          {
            from: env,
            to: '.'
          }
        ]
      })
    ]
};