/* eslint-disable no-undef */

const devCerts = require( "office-addin-dev-certs" );
const CopyWebpackPlugin = require( "copy-webpack-plugin" );
const HtmlWebpackPlugin = require( "html-webpack-plugin" );

const urlDev = "https://localhost:3000/";
// const urlProd = "https://outlook-office-addins.juanjoarranz.info/";     // OnPrem: IIS Server
// const urlProd = "https://outlook-6001.office.juanjoarranz.info/";       // OnPrem: Linux Node Server
const urlProd = "https://outlook-6001.azurewebsites.net/";                 // Cloud: Azure Node App Service

async function getHttpsOptions() {
    const httpsOptions = await devCerts.getHttpsServerOptions();
    return { cacert: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async ( env, options ) => {
    const dev = options.mode === "development";
    const config = {
        devtool: "source-map",
        optimization: {
            // The default minified version in production mode does not work
            minimize: false // force to not minify
        },
        entry: {
            polyfill: [ "core-js/stable", "regenerator-runtime/runtime" ],
            onsend: "./src/onsend/onsend.ts",
            dialog: "./src/dialog/dialog.ts",
        },
        output: {
            devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
            clean: true,
        },
        resolve: {
            extensions: [ ".ts", ".tsx", ".html", ".js" ],
        },
        module: {
            rules: [
                {
                    test: /\.ts$/,
                    exclude: /node_modules/,
                    use: {
                        loader: "babel-loader",
                        options: {
                            presets: [ "@babel/preset-typescript" ],
                        },
                    },
                },
                {
                    test: /\.tsx?$/,
                    exclude: /node_modules/,
                    use: "ts-loader",
                },
                {
                    test: /\.html$/,
                    exclude: /node_modules/,
                    use: "html-loader",
                },
                // {
                //     test: /\.(png|jpg|jpeg|gif|ico)$/,
                //     type: "asset/resource",
                //     generator: {
                //         filename: "assets/[name][ext][query]",
                //     },
                // },
            ],
        },
        plugins: [
            new HtmlWebpackPlugin( {
                filename: "onsend.html",
                template: "./src/onsend/onsend.html",
                chunks: [ "polyfill", "onsend" ],
            } ),
            new CopyWebpackPlugin( {
                patterns: [
                    // {
                    //     from: "assets/*",
                    //     to: "assets/[name][ext][query]",
                    // },
                    {
                        from: "manifest*.xml",
                        to: "[name]" + "[ext]",
                        transform( content ) {
                            return content.toString().replace( new RegExp( urlDev, "g" ), urlProd );

                            if ( dev ) {
                                return content;
                            } else {
                                return content.toString().replace( new RegExp( urlDev, "g" ), urlProd );
                            }
                        },
                    },
                    {
                        from: "./src/dialog/dialog.css",
                        to: "dialog.css",
                    },
                ],
            } ),
            new HtmlWebpackPlugin( {
                filename: "dialog.html",
                template: "./src/dialog/dialog.html",
                chunks: [ "polyfill", "dialog" ],
            } ),
        ],
        devServer: {
            headers: {
                "Access-Control-Allow-Origin": "*",
            },
            https: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
            port: process.env.npm_package_config_dev_server_port || 3000,
        },
    };

    return config;
};
