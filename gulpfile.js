/// Gulp configuration for Typescript, SASS and Live Reload

'use strict';

var gulp = require('gulp'),
    plumber = require('gulp-plumber'),
    del = require('del'),

    sass = require('gulp-sass'),
    
    // debug = require('gulp-debug'),

    typescript = require('gulp-typescript'),
    sourcemaps = require('gulp-sourcemaps'),
    tsConfigGlob = require('tsconfig-glob'),
    merge = require('merge2'),

    connect = require('gulp-connect'),
    config = require('./gulpfile.config.json'),
    packageConfig = require('./package.json'),
    errorHandler = function (error) {
        console.log(error);
        this.emit('end');
    };

var tsProject = typescript.createProject('./tsconfig.json', {
    sortdest: true
});

gulp.task('clean', function (done) {
    return del(config.server.root, done);
});

gulp.task('ref', function () {
    return tsConfigGlob();
});

gulp.task('copy:libs', ['clean'], function (done) {
    return Object.keys(packageConfig.overrides)
        .map(function (key) {
            try {
                var value = packageConfig.overrides[key];
                if (Array.isArray(value)) {
                    var files = value.map(function (filePath) {
                        console.log('Copied: ', filePath);
                        return config.lib.source + "/" + key + "/" + filePath;
                    });
                    return gulp.src(files)
                        .pipe(plumber(errorHandler))
                        .pipe(gulp.dest(config.lib.dest));
                } else {
                    var file = config.lib.source + "/" + key + "/" + value;
                    console.log('Copied: ', value);
                    return gulp.src(file)
                        .pipe(plumber(errorHandler))
                        .pipe(gulp.dest(config.lib.dest));
                }
            } catch (exception) {
                console.error('Failed to load package: ', key);
            }
        });
});

gulp.task('compile:sass', function () {
    return gulp.src(config.app.source + "/**/*.scss")
        .pipe(plumber(errorHandler))
        .pipe(sass())
        .pipe(gulp.dest(config.app.dest))
        .pipe(connect.reload());
});

gulp.task('compile:ts', ['ref'], function () {
    var tsResult = tsProject.src()
        .pipe(sourcemaps.init())
        .pipe(plumber(errorHandler))
        .pipe(typescript(tsProject));

    tsResult.js
        .pipe(sourcemaps.write('./'))
        .pipe(gulp.dest(config.app.dest))
        .pipe(connect.reload());
});

gulp.task('copy:misc', function () {
    gulp.src(config.app.source + "/**/!(*.ts|*.scss)", {
            base: config.app.source
        })
        .pipe(plumber(errorHandler))
        .pipe(gulp.dest(config.app.dest));
});

gulp.task('refresh', ['copy:misc'], function () {
    gulp.src(config.app.source + "/**/!(*.ts|*.scss)")
        .pipe(connect.reload())
        .pipe(plumber(errorHandler));
});

gulp.task('watch', function () {
    gulp.watch(config.app.source + "/**/*.scss", ['compile:sass']);
    gulp.watch(config.app.source + "/**/*.ts", ['compile:ts']);
    gulp.watch(config.app.source + "/**/!(*.ts|*.scss)", ['refresh']);
});

gulp.task('compile', ['compile:sass', 'compile:ts', 'copy:misc']);
gulp.task('build', ['generate-references', 'copy:libs']);
gulp.task('default', ['compile', 'watch'], function () {
    return connect.server({
        root: config.server.root,
        host: config.server.host,
        port: config.server.port,
        https: config.server.https,
        livereload: true,
        debug: true
    });
       
});