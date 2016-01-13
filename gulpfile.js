/// Gulp configuration for Live Reload

'use strict';

var gulp = require('gulp'),
    plumber = require('gulp-plumber'),
    connect = require('gulp-connect'),
    config = require('./gulpfile.config.json'),
    errorHandler = function (error) {
        console.log(error);
        this.emit('end');
    };

// Start the service.
gulp.task('serve', function () {
    connect.server({
        root: config.server.root,
        host: config.server.host,
        port: config.server.port,
        fallback: config.server.fallback,
        https: config.server.https,
        livereload: true
    });
});

// Reload the app when ever a file is changed.
gulp.task('server:reload', function () {
    gulp.src(config.app.source + "/**/*.*")
        .pipe(connect.reload())
        .pipe(plumber(errorHandler));
});

// Monitor the source folder (www) for any file changes and 
// then run the server:reload task.
gulp.task('watch', function () {
    gulp.watch(config.app.source + "/**/*.*", ['server:reload']);
});

gulp.task('default', ['serve', 'watch']);