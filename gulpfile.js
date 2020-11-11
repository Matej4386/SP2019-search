'use strict';
const os = require('os');
const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const path = require('path');
const log = require('@microsoft/gulp-core-build').log;
const colors = require("colors");
const bundleAnalyzer = require('webpack-bundle-analyzer');
const fs = require("fs");
/**
 * Task is necessary for VSTS
 */
build.addSuppression(/^Warning - \[sass\].*$/);

const envCheck = build.subTask('environmentCheck', (gulp, config, done) => {
    /********************************************************************************************
    * Adds an alias for handlebars in order to avoid errors while gulping the project
    * https://github.com/wycats/handlebars.js/issues/1174
    * Adds a loader and a node setting for webpacking the handlebars-helpers correctly
    * https://github.com/helpers/handlebars-helpers/issues/263
    ********************************************************************************************/
    build.configureWebpack.mergeConfig({
      additionalConfiguration: (generatedConfiguration) => {
        fs.writeFileSync("./temp/_webpack_config.json", JSON.stringify(generatedConfiguration, null, 2));
          /*
        generatedConfiguration.resolve.alias = {
          handlebars: 'handlebars/dist/handlebars.min.js'
        };*/

        // Check if running in debug or ship
        if (config.production) {
          generatedConfiguration.module.rules.push({
            test: /\.js$/,
            loader: 'unlazy-loader'
          });
        } else {
          generatedConfiguration.module.rules.push({
            test: /\.js$/,
            loader: 'unlazy-loader',
            exclude: [path.resolve(__dirname, './lib/')]
          });
        }


        generatedConfiguration.node = {
          fs: 'empty',
          readline: 'empty'
        }

        if (config.production) {
            log(`[${colors.cyan('configure-webpack')}] Adding plugin ${colors.cyan('BundleAnalyzerPlugin')}...`);
            const lastDirName = path.basename(__dirname);
            const dropPath = path.join(__dirname, 'temp', 'stats');
            generatedConfiguration.plugins.push(new bundleAnalyzer.BundleAnalyzerPlugin({
                openAnalyzer: false,
                analyzerMode: 'static',
                reportFilename: path.join(dropPath, `${lastDirName}.stats.html`),
                generateStatsFile: true,
                statsFilename: path.join(dropPath, `${lastDirName}.stats.json`),
                logLevel: 'error'
            }));
        }
        // Optimize build times - https://www.eliostruyf.com/speed-sharepoint-framework-builds-wsl-2/
        if (!config.production) {
            for (const rule of generatedConfiguration.module.rules) {
                // Add include rule for webpack's source map loader
                if (rule.use && typeof rule.use === 'string' && rule.use.indexOf('source-map-loader') !== -1) {
                    rule.include = [
                        path.resolve(__dirname, 'lib')
                    ]
                }

                // Disable minification for css-loader
                if (rule.use && rule.use instanceof Array && rule.use.length == 2 && rule.use[1].loader && rule.use[1].loader.indexOf('css-loader') !== -1) {
                    log(`[${colors.cyan('configure-webpack')}] Setting ${colors.cyan('css-loader')} to disable minification`);
                    rule.use[1].options.minimize = false;
                }
            }
        }

        return generatedConfiguration;
      }
    });

    done();
});
build.rig.addPreBuildTask(envCheck);

build.initialize(gulp);