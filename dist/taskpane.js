!function(e){var n={};function t(r){if(n[r])return n[r].exports;var o=n[r]={i:r,l:!1,exports:{}};return e[r].call(o.exports,o,o.exports,t),o.l=!0,o.exports}t.m=e,t.c=n,t.d=function(e,n,r){t.o(e,n)||Object.defineProperty(e,n,{enumerable:!0,get:r})},t.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},t.t=function(e,n){if(1&n&&(e=t(e)),8&n)return e;if(4&n&&"object"==typeof e&&e&&e.__esModule)return e;var r=Object.create(null);if(t.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:e}),2&n&&"string"!=typeof e)for(var o in e)t.d(r,o,function(n){return e[n]}.bind(null,o));return r},t.n=function(e){var n=e&&e.__esModule?function(){return e.default}:function(){return e};return t.d(n,"a",n),n},t.o=function(e,n){return Object.prototype.hasOwnProperty.call(e,n)},t.p="",t(t.s=308)}({308:function(e,n){function t(e,n,t,r,o,c,i){try{var u=e[c](i),a=u.value}catch(e){return void t(e)}u.done?n(a):Promise.resolve(a).then(r,o)}function r(e){return function(){var n=this,r=arguments;return new Promise((function(o,c){var i=e.apply(n,r);function u(e){t(i,o,c,u,a,"next",e)}function a(e){t(i,o,c,u,a,"throw",e)}u(void 0)}))}}function o(e){return c.apply(this,arguments)}function c(){return(c=r(regeneratorRuntime.mark((function e(n){return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.abrupt("return",Excel.run((function(e){return e.sync().then((function(){console.log("Change type of event: "+n.changeType),console.log("Address of event: "+n.address),console.log("Source of event: "+n.source),e.workbook.getSelectedRange().format.fill.color="yellow"}))})));case 1:case"end":return e.stop()}}),e)})))).apply(this,arguments)}Office.initialize=function(){Office.addin.getStartupBehavior().then((function(e){"Load"==e?(console.log("startupBehavior is set to Load"),$("#chk-set").prop("checked",!0)):(console.log("startupBehavior is set to not Load"),$("#chk-set").prop("checked",!1)),console.log(e)})),document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",Excel.run(function(){var e=r(regeneratorRuntime.mark((function e(n){return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return n.workbook.worksheets.getActiveWorksheet().onChanged.add(o),e.next=4,n.sync();case 4:console.log("A handler has been registered for the onChanged event.");case 5:case"end":return e.stop()}}),e)})));return function(n){return e.apply(this,arguments)}}())},$("#chk-set").on("change",(function(){this.checked?(console.log("Checked. Turning on startupBehavior."),Office.addin.setStartupBehavior(Office.StartupBehavior.load)):(console.log("Unchecked. Turning off startupBehavior"),Office.addin.setStartupBehavior(Office.StartupBehavior.none))}))}});
//# sourceMappingURL=taskpane.js.map