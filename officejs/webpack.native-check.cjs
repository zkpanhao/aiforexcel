const path = require('node:path');
const base = require('./webpack.config.js');
module.exports = async (env, argv) => {
  const config = await base(env, argv);
  config.entry = './src/excel/native-check.ts';
  config.output = { ...config.output, path: path.join(__dirname, 'build', 'native-check') };
  return config;
};
