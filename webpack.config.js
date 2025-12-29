const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, argv) => {
    const isDevelopment = argv.mode === 'development';

    return {
        entry: {
            taskpane: './src/taskpane/taskpane.tsx'
        },
        output: {
            path: path.resolve(__dirname, 'dist'),
            filename: '[name].js',
            clean: true
        },
        resolve: {
            extensions: ['.ts', '.tsx', '.js', '.jsx']
        },
        module: {
            rules: [
                {
                    test: /\.tsx?$/,
                    use: 'ts-loader',
                    exclude: /node_modules/
                },
                {
                    test: /\.css$/,
                    use: ['style-loader', 'css-loader']
                },
                {
                    test: /\.(png|jpg|jpeg|gif|ico)$/,
                    type: 'asset/resource'
                }
            ]
        },
        plugins: [
            new HtmlWebpackPlugin({
                template: './src/taskpane/taskpane.html',
                filename: 'taskpane.html',
                chunks: ['taskpane']
            }),
            new CopyWebpackPlugin({
                patterns: [
                    { from: 'assets', to: 'assets', noErrorOnMissing: true }
                ]
            })
        ],
        devServer: {
            static: {
                directory: path.join(__dirname, 'dist')
            },
            headers: {
                'Access-Control-Allow-Origin': '*'
            },
            port: 3000,
            hot: true,
            server: 'https'
        },
        devtool: isDevelopment ? 'source-map' : false
    };
};
