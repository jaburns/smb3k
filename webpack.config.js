const webpack = require('webpack');

module.exports = {
    context: __dirname,
    devtool: 'source-map',
    mode: 'production',
    entry: './js/engine.js',
    output: {
        path: __dirname+'/public',
        filename: 'bundle.js'
    },
    module: {
        rules: [{ 
            test: /\.js$/, exclude: /node_modules/,
            use: {
                loader: 'babel-loader', 
                options: {presets: ['babel-preset-env']}
            }
        }]
    },
};