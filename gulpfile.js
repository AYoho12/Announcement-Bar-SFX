'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(100,45): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(104,20): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(222,82): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(108,44): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(135,78): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(135,93): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(225,19): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
build.addSuppression(`Warning - lint - src/webparts/announcementBanner/AnnouncementBannerWebPart.ts(108,46): error @typescript-eslint/no-explicit-any: Unexpected any. Specify a different type.`);
var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
