'use strict';

// This runs 'hello-world/node_modules/@microsoft/sp-build-web/lib-commonjs/index.js' file
const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;

// It overrides the getTasks function to use serve-deprecated instead of serve
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
