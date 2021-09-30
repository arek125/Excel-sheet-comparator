const path = require('path');

module.exports = {
    entry: './preload.js',
    output: {
      filename: 'main.js',
      path: path.resolve(__dirname, 'dist'),
    },
    resolve: {
        extensions: ['.ts', '.js'],
    },
    target: ['web', 'es5'],
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: /(node_modules|bower_components)/,
                use: {
                    loader: 'babel-loader',
                    options: {
                        babelrc: false,
                        configFile: path.resolve(__dirname, 'babel.config.js'),
                        compact: false,
                        cacheDirectory: true,
                        sourceMaps: false,
                    },
                }
            }
        ]
    },
};