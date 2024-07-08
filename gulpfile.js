'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(102,45): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(111,78): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(111,93): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(165,80): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(168,17): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
