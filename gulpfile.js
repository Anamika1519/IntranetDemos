process.on('unhandledRejection', error => {
  console.error('🛑 Unhandled Rejection:', error);
});
process.on('uncaughtException', error => {
  console.error('🛑 Uncaught Exception:', error);
});
'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};
process.on('unhandledRejection', (reason) => {
  console.error('🛑 Unhandled Rejection:\n', reason?.stack || JSON.stringify(reason, null, 2));
});

process.on('uncaughtException', (error) => {
  console.error('🛑 Uncaught Exception:\n', error?.stack || JSON.stringify(error, null, 2));
});
build.initialize(require('gulp'));
