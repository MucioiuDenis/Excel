!function(){"use strict";var e={14183:function(e,t){var n=this&&this.__awaiter||function(e,t,n,r){return new(n||(n=Promise))((function(o,i){function c(e){try{a(r.next(e))}catch(e){i(e)}}function s(e){try{a(r.throw(e))}catch(e){i(e)}}function a(e){var t;e.done?o(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(c,s)}a((r=r.apply(e,t||[])).next())}))},r=this&&this.__generator||function(e,t){var n,r,o,i,c={label:0,sent:function(){if(1&o[0])throw o[1];return o[1]},trys:[],ops:[]};return i={next:s(0),throw:s(1),return:s(2)},"function"==typeof Symbol&&(i[Symbol.iterator]=function(){return this}),i;function s(s){return function(a){return function(s){if(n)throw new TypeError("Generator is already executing.");for(;i&&(i=0,s[0]&&(c=0)),c;)try{if(n=1,r&&(o=2&s[0]?r.return:s[0]?r.throw||((o=r.return)&&o.call(r),0):r.next)&&!(o=o.call(r,s[1])).done)return o;switch(r=0,o&&(s=[2&s[0],o.value]),s[0]){case 0:case 1:o=s;break;case 4:return c.label++,{value:s[1],done:!1};case 5:c.label++,r=s[1],s=[0];continue;case 7:s=c.ops.pop(),c.trys.pop();continue;default:if(!((o=(o=c.trys).length>0&&o[o.length-1])||6!==s[0]&&2!==s[0])){c=0;continue}if(3===s[0]&&(!o||s[1]>o[0]&&s[1]<o[3])){c.label=s[1];break}if(6===s[0]&&c.label<o[1]){c.label=o[1],o=s;break}if(o&&c.label<o[2]){c.label=o[2],c.ops.push(s);break}o[2]&&c.ops.pop(),c.trys.pop();continue}s=t.call(e,c)}catch(e){s=[6,e],r=0}finally{n=o=0}if(5&s[0])throw s[1];return{value:s[0]?s[1]:void 0,done:!0}}([s,a])}}};function o(){return n(this,void 0,void 0,(function(){var e,t=this;return r(this,(function(o){switch(o.label){case 0:return o.trys.push([0,2,,3]),[4,Excel.run((function(e){return n(t,void 0,void 0,(function(){var t;return r(this,(function(n){switch(n.label){case 0:return(t=e.workbook.getSelectedRange()).load("address"),t.format.fill.color="red",[4,e.sync()];case 1:return n.sent(),console.log("The range address was ".concat(t.address,".")),[2]}}))}))}))];case 1:return o.sent(),[3,3];case 2:return e=o.sent(),console.error(e),[3,3];case 3:return[2]}}))}))}function i(){return n(this,void 0,void 0,(function(){var e,t=this;return r(this,(function(o){switch(o.label){case 0:return o.trys.push([0,2,,3]),[4,Excel.run((function(e){return n(t,void 0,void 0,(function(){var t,n;return r(this,(function(r){switch(r.label){case 0:return t=e.workbook.worksheets.getItem("Sheet1"),(n=t.tables.add("A1:D1",!0)).name="ExpensesTable",n.getHeaderRowRange().values=[["Date","Merchant","Category","Amount"]],n.rows.add(null,[["1/1/2017","The Phone Company","Communications","$120"],["1/2/2017","Northwind Electric Cars","Transportation","$142"],["1/5/2017","Best For You Organics Company","Groceries","$27"],["1/10/2017","Coho Vineyard","Restaurant","$33"],["1/11/2017","Bellows College","Education","$350"],["1/15/2017","Trey Research","Other","$135"],["1/15/2017","Best For You Organics Company","Groceries","$97"]]),Office.context.requirements.isSetSupported("ExcelApi","1.2")&&(t.getUsedRange().format.autofitColumns(),t.getUsedRange().format.autofitRows()),t.activate(),[4,e.sync()];case 1:return r.sent(),[2]}}))}))}))];case 1:return o.sent(),[3,3];case 2:return e=o.sent(),console.error(e),[3,3];case 3:return[2]}}))}))}function c(){return n(this,void 0,void 0,(function(){var e,t=this;return r(this,(function(o){switch(o.label){case 0:return o.trys.push([0,2,,3]),[4,Excel.run((function(e){return n(t,void 0,void 0,(function(){var t,n;return r(this,(function(r){switch(r.label){case 0:return t=e.workbook.worksheets.getItem("Sheet1").shapes,(n=t.addGeometricShape(Excel.GeometricShapeType.rectangle)).left=100,n.top=100,n.height=150,n.width=150,n.name="Square",[4,e.sync()];case 1:return r.sent(),[2]}}))}))}))];case 1:return o.sent(),[3,3];case 2:return e=o.sent(),console.error(e),[3,3];case 3:return[2]}}))}))}function s(){return n(this,void 0,void 0,(function(){var e,t=this;return r(this,(function(o){switch(o.label){case 0:return o.trys.push([0,2,,3]),[4,Excel.run((function(e){return n(t,void 0,void 0,(function(){return r(this,(function(t){switch(t.label){case 0:return e.workbook.worksheets.getItem("Sheet1").shapes.addLine(200,50,300,150,Excel.ConnectorType.straight).name="StraightLine",[4,e.sync()];case 1:return t.sent(),[2]}}))}))}))];case 1:return o.sent(),[3,3];case 2:return e=o.sent(),console.error(e),[3,3];case 3:return[2]}}))}))}Object.defineProperty(t,"__esModule",{value:!0}),t.createPowerPoint=t.line=t.shape=t.table=t.run=void 0,Office.onReady((function(){document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("run").onclick=o,document.getElementById("table").onclick=i,document.getElementById("square").onclick=c,document.getElementById("line").onclick=s})),t.run=o,t.table=i,t.shape=c,t.line=s,t.createPowerPoint=function(){return n(this,void 0,void 0,(function(){return r(this,(function(e){try{new Object(PowerPoint.Application)}catch(e){console.error(e)}return[2]}))}))}},93823:function(e,t,n){var r=n(27091),o=n.n(r),i=new URL(n(60806),n.b),c=new URL(n(44944),n.b);o()(i),o()(c)},27091:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},44944:function(e,t,n){e.exports=n.p+"assets/logo-filled.png"},60806:function(e,t,n){e.exports=n.p+"1fda685b81e1123773f6.css"}},t={};function n(r){var o=t[r];if(void 0!==o)return o.exports;var i=t[r]={exports:{}};return e[r].call(i.exports,i,i.exports,n),i.exports}n.m=e,n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,{a:t}),t},n.d=function(e,t){for(var r in t)n.o(t,r)&&!n.o(e,r)&&Object.defineProperty(e,r,{enumerable:!0,get:t[r]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;n.g.importScripts&&(e=n.g.location+"");var t=n.g.document;if(!e&&t&&(t.currentScript&&(e=t.currentScript.src),!e)){var r=t.getElementsByTagName("script");r.length&&(e=r[r.length-1].src)}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=e}(),n.b=document.baseURI||self.location.href,n(14183),n(93823)}();
//# sourceMappingURL=taskpane.js.map