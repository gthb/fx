"use strict";const e="operator",t="error",n="range_beam",l="range_ternary",r="range_named",o="structured",u="unknown",s="UnaryExpression",c="BinaryExpression",i="ReferenceIdentifier",a="CallExpression";function f(e){let t=arguments.length>1&&void 0!==arguments[1]&&arguments[1],n=!1,l="";const r=[],o=()=>{l&&r.push(t?l:{value:l,braced:n}),l=""};for(let t=0;t<e.length;t++){const r=e[t];"["===r?(o(),n=!0):"]"===r?(o(),n=!1):l+=r}return o(),r}function p(e){return{context:f(e,!0)}}function h(e){const t={},n=f(e);if(n.length>1)t.workbookName=n[n.length-2].value,t.sheetName=n[n.length-1].value;else if(1===n.length){const e=n[0];e.braced?t.workbookName=e.value:t.sheetName=e.value}return t}const g=e=>e&&":"===e.value&&{},x=e=>e&&"range"===e.type&&{r0:e.value},d=e=>e&&e.type===l&&{r0:e.value},$=e=>e&&"range"===e.type&&{r1:e.value},y=t=>t&&t.type===e&&"!"===t.value&&{},m=e=>e&&e.type===n&&{r0:e.value},v=e=>e&&e.type===o&&{struct:e.value},E=(e,t)=>{const n=t.xlsx?h:p;return e&&"context"===e.type?n(e.value):e&&"context_quote"===e.type?n(e.value.slice(1,-1).replace(/''/g,"'")):void 0},R=e=>e&&e.type===r&&{name:e.value},w=[[d],[x,g,$],[x],[m],[E,y,d],[E,y,x,g,$],[E,y,x],[E,y,m]],N=w.concat([[R],[E,y,R],[v],[R,v],[E,y,R,v]]);function b(e,t){const n={withLocation:!1,mergeRefs:!1,allowTernary:!1,allowNamed:!0,r1c1:!1,xlsx:!1,...t},l=Fe(e,Ae,n),r=n.xlsx?{workbookName:"",sheetName:"",r0:"",r1:"",name:""}:{context:[],r0:"",r1:"",name:""};l.length&&"fx_prefix"===l[0].type&&l.shift();const o=n.allowNamed?N:w;for(let e=0;e<o.length;e++){const t={...r};if(o[e].length===l.length){const r=o[e].every(((e,r)=>{const o=e(l[r],n);return Object.assign(t,o),o}));if(r)return t}}return null}const A=/[^0-9A-Za-z._¡¤§¨ª\u00ad¯-\uffff]/;function C(e){let t="",n=0,l=0;const r=e.context||[];for(let e=r.length;e>-1;e--){const o=r[e];if(o){t=(l%2?"["+o+"]":o)+t,n+=A.test(o),l++}}return n&&(t="'"+t.replace(/'/g,"''")+"'"),t?t+"!":t}function T(e){let t="",n=0;const{workbookName:l,sheetName:r}=e;return l&&(t+="["+l+"]",n+=A.test(l)),r&&(t+=r,n+=A.test(r)),n&&(t="'"+t.replace(/'/g,"''")+"'"),t?t+"!":t}const I=(e,t,n)=>Math.min(Math.max(t,e),n),L=(e,t)=>(t?"$":"")+U(e),k=(e,t)=>(t?"$":"")+String(e+1),O=String.fromCharCode;function _(e){const t=e||"",n=t.length;let l=0;if(n>2){const e=t.charCodeAt(n-3);l+=676*(1+e-(e>95?32:0)-65)}if(n>1){const e=t.charCodeAt(n-2);l+=26*(1+e-(e>95?32:0)-65)}if(n){const e=t.charCodeAt(n-1);l+=e-(e>95?32:0)-65}return l}function U(e){return(e>=702?O(((e-702)/676-0)%26+65):"")+(e>=26?O((e/26-1)%26+65):"")+O(e%26+65)}function F(e){let{top:t,left:n,bottom:l,right:r}=e;const{$left:o,$right:u,$top:s,$bottom:c}=e,i=null==n,a=null==r,f=null==t,p=null==l;t=I(0,0|t,1048575),n=I(0,0|n,16383),!i&&!f&&a&&p?(l=t,r=n):(l=I(0,0|l,1048575),r=I(0,0|r,16383));if(0===t&&l>=1048575&&!i&&!a&&(!(o&&!i||u&&!a)||n===r)||f&&p)return L(n,o)+":"+L(r,u);return 0===n&&r>=16383&&!f&&!p&&(!(s&&!f||c&&!p)||t===l)||i&&a?k(t,s)+":"+k(l,c):i||f||a||!p?i||!f||a||p?i||f||!a||p?!i||f||a||p?r!==n||l!==t||u!==o||c!==s?L(n,o)+k(t,s)+":"+L(r,u)+k(l,c):L(n,o)+k(t,s):L(r,u)+k(t,s)+":"+k(l,c):L(n,o)+k(t,s)+":"+k(l,c):L(n,o)+k(l,c)+":"+L(r,u):L(n,o)+k(t,s)+":"+L(r,u)}function S(e){const t=/^(?=.)(\$(?=\D))?([A-Za-z]{0,3})?(\$)?([1-9][0-9]{0,6})?$/.exec(e);return t&&(t[2]||t[4])?[t[4]?(n=t[4],+n-1):null,t[2]?_(t[2]):null,!!t[3],!!t[1]]:null;var n}function M(e){let t=null,n=null,l=null,r=null,o=!1,u=!1,s=!1,c=!1;const[i,a,f]=e.split(":");if(f)return null;const p=S(i),h=a?S(a):null;if(!p||a&&!h)return null;if(null!=p[0]&&null!=p[1]?[t,n,o,u]=p:null==p[0]&&null!=p[1]?[,n,,u]=p:null!=p[0]&&null==p[1]&&([t,,o]=p),a)null!=h[0]&&null!=h[1]?[l,r,s,c]=h:null==h[0]&&null!=h[1]?[,r,,c]=h:null!=h[0]&&null==h[1]&&([l,,s]=h);else{if(null==t||null==n)return null;l=t,r=n,s=o,c=u}return null!=r&&(null==n||null!=n&&r<n)&&([n,r,u,c]=[r,n,c,u]),null!=l&&(null==t||null!=t&&l<t)&&([t,l,o,s]=[l,t,s,o]),{top:t,left:n,bottom:l,right:r,$top:o,$left:u,$bottom:s,$right:c}}function z(e){let{allowNamed:t=!0,allowTernary:n=!1,xlsx:l=!1}=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};const r=b(e,{allowNamed:t,allowTernary:n,xlsx:l,r1c1:!1});if(r&&(r.r0||r.name)){let e=null;return r.r0&&(e=M(r.r1?r.r0+":"+r.r1:r.r0)),r.name||e?(r.range=e,delete r.r0,delete r.r1,r):null}return null}function D(e){let{xlsx:t=!1}=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};const n=t?T(e):C(e);return n+(e.name?e.name:F(e.range))}function j(e){return null==e.top&&(e.top=0,e.$top=!1),null==e.bottom&&(e.bottom=1048575,e.$bottom=!1),null==e.left&&(e.left=0,e.$left=!1),null==e.right&&(e.right=16383,e.$right=!1),e}const Z=/^\[('['#@[\]]|[^'#@[\]])+\]/i,q=/^([^#@[\]:]+)/i,B={headers:1,data:2,totals:4,all:8,"this row":16,"@":16},P=function(){for(var e=arguments.length,t=new Array(e),n=0;n<e;n++)t[n]=arguments[n];return Object.freeze(t)},X={0:P(),1:P("headers"),2:P("data"),4:P("totals"),8:P("all"),16:P("this row"),3:P("headers","data"),6:P("data","totals")},W=function(e){let t=!(arguments.length>1&&void 0!==arguments[1])||arguments[1],n=Z.exec(e);if(n){const e=n[0].slice(1,-1).replace(/'(['#@[\]])/g,"$1");return[n[0],e]}return t&&(n=q.exec(e),n)?[n[0],n[0]]:null};function H(e){const t=[];let n,l,r=0,o=e,u=0;if(!(n=/^(\[\s*)/.exec(o)))return null;if(l=/^\[#([a-z ]+)\]/i.exec(o)){const e=l[1].toLowerCase();if(r+=l[0].length,!B[e])return null;u|=B[e]}else if(l=W(o,!1))r+=l[0].length,t.push(l[1].trim());else{let l=!0;for(o=o.slice(n[1].length),r+=n[1].length;l&&(n=/^\[#([a-z ]+)\](\s*,\s*)?/i.exec(o));){const e=n[1].toLowerCase();if(!B[e])return null;u|=B[e],o=o.slice(n[0].length),r+=n[0].length,l=!!n[2]}if(l&&(n=/^@/.exec(o))&&(u|=B["@"],o=o.slice(1),r+=1,l="]"!==o[0]),!(u in X))return null;const s=l?W(e.slice(r)):null;if(s){if(r+=s[0].length,t.push(s[1].trim()),o=e.slice(r),":"===o[0]){o=o.slice(1),r++;const e=W(o);if(!e)return null;r+=e[0].length,t.push(e[1].trim())}l=!1}for(;" "===e[r];)r++;if(l||"]"!==e[r])return null;r++}const s=X[u];return{columns:t,sections:s?s.concat():s,length:r,token:e.slice(0,r)}}function Y(e){let t=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{xlsx:!1};const n=b(e,t);if(n&&n.struct){const e=H(n.struct);if(e&&e.length===n.struct.length)return t.xlsx?{workbookName:n.workbookName,sheetName:n.sheetName,table:n.name,columns:e.columns,sections:e.sections}:{context:n.context,table:n.name,columns:e.columns,sections:e.sections}}return null}function G(e){return e.replace(/([[\]#'@])/g,"'$1")}function K(e){return!/^[a-zA-Z0-9\u00a1-\uffff]+$/.test(e)}function V(e){return e[0].toUpperCase()+e.slice(1).toLowerCase()}function Q(e){let{xlsx:t=!1}=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{},n=t?T(e):C(e);e.table&&(n+=e.table);const l=e.columns?.length??0,r=e.sections?.length??0;if(1!==r||l)if(r||1!==l){n+="[";const t=1===r&&"this row"===e.sections[0].toLowerCase();t?n+="@":r&&(n+=e.sections.map((e=>`[#${V(e)}]`)).join(","),l&&(n+=",")),t&&1===e.columns.length&&!K(e.columns[0])?n+=G(e.columns[0]):l&&(n+=e.columns.slice(0,2).map((e=>`[${G(e)}]`)).join(":")),n+="]"}else n+=`[${G(e.columns[0])}]`;else n+=`[#${V(e.sections[0])}]`;return n}const J=/^(?!!)(\[(?:[^\]])+\])?([0-9A-Za-z._¡¤§¨ª\u00ad¯-\uffff]+)?(?=!)/,ee=/^'(?:''|[^'])*('|$)(?=!)/,te="\\$?[A-Z]{1,3}\\$?[1-9][0-9]{0,6}",ne="\\$?[A-Z]{1,3}",le="\\$?[1-9][0-9]{0,6}",re="(?![a-z0-9_\\u00a1-\\uffff])",oe=new RegExp(`^${ne}:${ne}${re}`,"i"),ue=new RegExp(`^${le}:${le}${re}`,"i"),se=new RegExp(`^${te}${re}`,"i"),ce=new RegExp(`^((${ne}|${le}):${te}|${te}:(${ne}|${le}))(?![\\w($.])`,"i"),ie="(?:R(?:\\[[+-]?\\d+\\]|[1-9][0-9]{0,6})?)",ae="(?:C(?:\\[[+-]?\\d+\\]|[1-9][0-9]{0,4})?)",fe=new RegExp(`^${ae}(:${ae})?${re}`,"i"),pe=new RegExp(`^${ie}(:${ie})?${re}`,"i"),he=new RegExp(`^(?:(?=[RC])${ie}${ae})${re}`,"i"),ge=new RegExp(`^(${ie}${ae}(:${ae}|:${ie})(?![[\\d])|(${ie}|${ae})(:${ie}${ae}))${re}`,"i"),xe=/^[a-zA-Z\\_\u00a1-\uffff][a-zA-Z0-9\\_.?\u00a1-\uffff]{0,254}/i;function de(e,t){return n=>{const l=t.exec(n);if(l)return{type:e,value:l[0]}}}function $e(e){const t=xe.exec(e);if(t){const e=t[0].toLowerCase();return"r"===e||"c"===e?null:{type:r,value:t[0]}}}const ye=/^'(?:[^[\]]+?)?(?:\[(.+?)\])?(?:[^[\]]+?)'$/,me=/^'\[(.+?)\]'$/;function ve(e,t){const n=ee.exec(e);if(n){const e=n[0];if(t.xlsx&&me.test(e)||ye.test(e))return{type:"context_quote",value:e}}const l=J.exec(e);if(l){const[,e,n]=l;if(e&&n||n||e&&!n&&t.xlsx)return{type:"context",value:l[0]}}}function Ee(e){const t=H(e);if(t){let n=t.length;for(;" "===e[n];)n++;if("!"!==e[n])return{type:o,value:t.token}}return null}const Re=/([RC])(\[?)(-?\d+)/gi,we=/(\d+|[a-zA-Z]+)/gi;function Ne(e,t){let r,o;if(t.r1c1){if(t.allowTernary&&(r=ge.exec(e))?o={type:l,value:r[0]}:(r=he.exec(e))?o={type:"range",value:r[0]}:((r=pe.exec(e))||(r=fe.exec(e)))&&(o={type:n,value:r[0]}),o){for(Re.lastIndex=0;null!==(r=Re.exec(o.value));){const e=("R"===r[1]?1048575:16383)+(r[2]?0:1),t=parseInt(r[3],10);if(t>e||t<-e)return null}return o}}else if(t.allowTernary&&(r=ce.exec(e))?o={type:l,value:r[0]}:(r=oe.exec(e))||(r=ue.exec(e))?o={type:n,value:r[0]}:(r=se.exec(e))&&(o={type:"range",value:r[0]}),o){for(we.lastIndex=0;null!==(r=we.exec(o.value));)if(/^\d/.test(r[1])){if(parseInt(r[1],10)-1>1048575)return null}else if(_(r[1])>16383)return null;return o}}const be=[de(t,/^#(NAME\?|FIELD!|CALC!|VALUE!|REF!|DIV\/0!|NULL!|NUM!|N\/A|GETTING_DATA\b|SPILL!|UNKNOWN!|FIELD\b|CALC\b|SYNTAX\?|ERROR!|CONNECT!|BLOCKED!|EXTERNAL!)/i),de(e,/^(<=|>=|<>|[-+/*^%&<>=]|[{},;]|[()]|@|:|!|#)/),de("func",/^[A-Z_]+[A-Z\d_.]*(?=\()/i),de("bool",/^(TRUE|FALSE)\b/i),de("newline",/^\n+/),de("whitespace",/^[ \f\r\t\v\u00a0\u1680\u2000-\u200a\u2028\u2029\u202f\u205f\u3000\ufeff]+/),de("string",/^"(?:""|[^"])*("|$)/),ve,Ne,Ee,de("number",/^(?:\d+(\.\d+)?(?:[eE][+-]?\d+)?|\d+)/),$e],Ae=[function(t,n){return n.r1c1?"!"===t[0]?{type:e,value:t[0]}:null:"!"===t[0]||":"===t[0]?{type:e,value:t[0]}:null},ve,Ne,Ee,$e],Ce={};function Te(e,t){if(e.length){const n=e[0];t[n]=t[n]||{},Te(e.slice(1),t[n])}else t.$=!0}[["range",":","range"],["range"],[n],[l],["context","!","range",":","range"],["context","!","range"],["context","!",n],["context","!",l],["context_quote","!","range",":","range"],["context_quote","!","range"],["context_quote","!",n],["context_quote","!",l],[r],["context","!",r],["context_quote","!",r],[o],[r,o],["context","!",r,o],["context_quote","!",r,o]].forEach((e=>Te(e.concat().reverse(),Ce)));const Ie=function(t,n,l){let r=arguments.length>3&&void 0!==arguments[3]?arguments[3]:0;const o=t[l-r];if(o){const u=o.type===e?o.value:o.type;if(u in n)return Ie(t,n[u],l,r+1)}return n.$?r:0};function Le(e){const t=[];for(let n=e.length-1;n>=0;n--){let l=e[n];const r=Ie(e,Ce,n);if(r){const t=e.slice(n-r+1,n+1);l={...l},l.value=t.map((e=>e.value)).join(""),l.loc&&t[0].loc&&(l.loc[0]=t[0].loc[0]),n-=r-1}t.unshift(l)}return t}const ke=(e,t)=>e&&e.type===t,Oe={withLocation:!1,mergeRefs:!0,allowTernary:!1,negativeNumbers:!0,r1c1:!1},_e=e=>e.type===r||"func"===e.type,Ue=t=>!ke(t,e)||"%"===t.value||"}"===t.value||")"===t.value||"#"===t.value;function Fe(t,n){let l=arguments.length>2&&void 0!==arguments[2]?arguments[2]:{};const r=Object.assign({},Oe,l),{withLocation:o,mergeRefs:s,negativeNumbers:c}=r,i=[];let a=0,f=null,p=null,h=null;const g=e=>{const t=e.type===u,n=h&&h.type===u;h&&(t&&n||t&&_e(h)||n&&_e(e))?(h.value+=e.value,h.type=u,o&&(h.loc[1]=e.loc[1])):(i.push(e),h=e,"whitespace"!==e.type&&"newline"!==e.type&&(p=f,f=e))};if(/^=/.test(t)){a++,g({type:"fx_prefix",value:"=",...o?{loc:[0,1]}:{}})}for(;a<t.length;){const l=a,s=t.slice(a);let x="",d="";for(let e=0;e<n.length;e++){const t=n[e](s,r);if(t){x=t.type,d=t.value,a+=d.length;break}}x||(x=u,d=t[a],a++);const $={type:x,value:d,...o?{loc:[l,a]}:{}};if("string"===x){const e=d.length;if('""'===d);else if('"'===d||'"'!==d[e-1])$.unterminated=!0;else if('""'!==d&&'"'===d[e-2]){let t=e-1;for(;'"'===d[t];)t--;!(t+1)^(e-t+1)%2==0&&($.unterminated=!0)}}if(c&&"number"===x){const t=h;if(t&&ke(t,e)&&"-"===t.value&&(!p||ke(p,"fx_prefix")||!Ue(p))){const e=i.pop();$.value="-"+d,o&&($.loc[0]=e.loc[0]),f=p,h=i[i.length-1]}}g($)}return s?Le(i):i}function Se(e){let t=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};return Fe(e,be,t)}function Me(e){return!!e&&("range"===e.type||e.type===n||e.type===l)}function ze(e){return!!e&&("range"===e.type||e.type===n||e.type===l||e.type===o||e.type===r)}function De(e){return!!e&&("bool"===e.type||e.type===t||"number"===e.type||"string"===e.type)}function je(e){return!!e&&e.type===t}function Ze(e){return!!e&&("whitespace"===e.type||"newline"===e.type)}function qe(e){return!!e&&"func"===e.type}function Be(e){return!!e&&"fx_prefix"===e.type}function Pe(t){return!!t&&t.type===e}const Xe="(END)",We=["ANCHORARRAY","CHOOSE","DROP","IF","IFS","INDEX","INDIRECT","LAMBDA","LET","OFFSET","REDUCE","SINGLE","SWITCH","TAKE","XLOOKUP"],He=e=>!!e&&(e.type===i||("ErrorLiteral"===e.type||e.type===t)&&"#REF!"===e.value||e.type===c&&(":"===e.operator||" "===e.operator||","===e.operator)||e.type===a&&We.includes(e.callee.name.toUpperCase())),Ye={};let Ge,Ke,Ve,Qe=!1,Je=!1;function et(e){const t=new Error(e);throw t.source=Ke.map((e=>e.value)).join(""),t.sourceOffset=Ke.slice(0,Ve).reduce(((e,t)=>e+t.value),"").length,t}function tt(){let e,t=Ve;do{e=Ke[++t]}while(e&&(Ze(e)||Pe(e)&&"("===e.value));return(e=>{const t=(e&&e.value)+"";return!(!ze(e)&&(!Pe(e)||":"!==t&&","!==t&&t.trim())&&(!qe(e)||!We.includes(t.toUpperCase()))&&(!je(e)||"#REF!"!==t))})(e)}function nt(){let e=arguments.length>0&&void 0!==arguments[0]?arguments[0]:null,t=arguments.length>1&&void 0!==arguments[1]?arguments[1]:null;if(e&&e!==Ge.id&&et(`Expected ${e} but got ${Ge.id}`),Ze(Ke[Ve])){if(!(He(t)&&tt()))for(;Ze(Ke[Ve]);)Ve++}if(Ve>=Ke.length)return void(Ge=Ye[Xe]);const n=Ke[Ve];let l;Ve+=1,n.unterminated&&et("Encountered an unterminated token");let r=n.type;return Pe(n)?(l=Ye[n.value],l||et(`Unknown operator ${n.value}`)):Ze(n)?l=Ye["(WHITESPACE)"]:De(n)?l=Ye.Literal:ze(n)?(l=Ye[i],r=i):qe(n)?l=Ye["(FUNCTION)"]:et(`Unexpected ${n.type} token: ${n.value}`),Ge=Object.create(l),Ge.type=r,Ge.value=n.value,n.loc&&(Ge.loc=[...n.loc]),Ge}function lt(e){let t=Ge;nt(null,t);let n=t.nud();for(;e<Ge.lbp;)t=Ge,nt(null,t),n=t.led(n);return n}const rt={nud:()=>et("Invalid syntax"),led:()=>et("Missing operator")};function ot(e){let t=arguments.length>1&&void 0!==arguments[1]?arguments[1]:0,n=Ye[e];return n?t>=n.lbp&&(n.lbp=t):(n={...rt},n.id=e,n.value=e,n.lbp=t,Ye[e]=n),n}function ut(e,t,n){const l=ot(e,t);return l.led=n||function(e){this.type=c,this.operator=this.value,delete this.value;const n=lt(t);return this.arguments=[e,n],this.loc&&(this.loc=[e.loc[0],n.loc[1]]),this},l}function st(e,t){const n=ot(e,0);return n.lbp=70,n.led=t||function(e){return this.type=s,this.operator=this.value,delete this.value,this.arguments=[e],this.loc&&(this.loc[0]=e.loc[0]),this},n}function ct(e,t){const n=ot(e);return n.nud=t||function(){this.type=s,this.operator=this.value,delete this.value;const e=lt(70);return this.arguments=[e],this.loc&&(this.loc[1]=e.loc[1]),this},n}function it(e,t){return ut(e,t,(function(n){He(n)||et(`Unexpected ${e} operator`);const l=lt(t);return He(l)||et(`Unexpected ${Ge.type} following ${this.id}`),this.type=c,this.operator=this.value.trim()?this.value:" ",delete this.value,this.arguments=[n,l],this.loc&&(this.loc=[n.loc[0],l.loc[1]]),this}))}ot(Xe),it(":",80);const at=it(",",80);it("(WHITESPACE)",80);const ft=e=>{const t=at.lbp>0;return null!=e&&(at.lbp=e?80:0),t};function pt(){let e=1;return()=>"fxg"+e++}function ht(e,t){return null==e&&null==t||e===t}function gt(e,t){if(Array.isArray(e)!==Array.isArray(t)||e.length!==t.length)return!1;for(let n=0;n<e.length;n++)if(!ht(e[n],t[n]))return!1;return!0}function xt(e,t){return!e&&!t||String(e).toLowerCase()===String(t).toLowerCase()}function dt(e,t){if((e.name||t.name)&&e.name!==t.name)return!1;if(e.columns||t.columns){if(e.table!==t.table)return!1;if(!gt(e.columns,t.columns))return!1;if(!gt(e.sections,t.sections))return!1}return!!(!e.range&&!t.range||ht(e.range.top,t.range.top)&&ht(e.range.bottom,t.range.bottom)&&ht(e.range.left,t.range.left)&&ht(e.range.right,t.range.right))&&!(!xt(e.workbookName,t.workbookName)||!xt(e.sheetName,t.sheetName))}function $t(e,t,n){return e.sheetName||(e.sheetName=t),e.workbookName||(e.workbookName=n),e}st("%"),st("#",(function(e){return He(e)||et("# expects a reference"),this.type=s,this.operator=this.value,delete this.value,this.arguments=[e],this})),ct("+"),ct("-"),ct("@"),ut("^",50),ut("*",40),ut("/",40),ut("+",30),ut("-",30),ut("&",20),ut("=",10),ut("<",10),ut(">",10),ut("<=",10),ut(">=",10),ut("<>",10),ot("Literal").nud=function(){const{type:e,value:n}=this;if(this.type="Literal",this.raw=n,"number"===e)this.value=+n;else if("bool"===e)this.value="TRUE"===n.toUpperCase();else if(e===t)this.type="ErrorLiteral",this.value=n.toUpperCase();else{if("string"!==e)throw new Error("Unsupported literal type: "+e);this.value=n.slice(1,-1).replace(/""/g,'"')}return this},ot(i).nud=function(){return this.type=i,this},ot(")"),ct("(",(function(){const e=ft(!0),t=lt(0);return nt(")",t),ft(e),t})),ot("(FUNCTION)").nud=function(){return this},ut("(",90,(function(e){"(FUNCTION)"!==e.id&&et("Cannot call a "+e.type);const t=[];let n=!1;if(")"!==Ge.id){const e=ft(!1);for(;")"!==Ge.id;)if(Ze(Ge)&&nt(),","===Ge.id)t.push(null),n=!0,nt();else{const e=lt(0);t.push(e),n=!1,","===Ge.id&&(nt(","),n=!0)}ft(e)}n&&t.push(null);const l=Ge;return delete this.value,this.type=a,this.callee={type:"Identifier",name:e.value},e.loc&&(this.callee.loc=[...e.loc]),this.arguments=t,e.loc&&(this.loc=[e.loc[0],l.loc[1]]),nt(")",this),this})),ot("}"),ot(";"),ct("{",(function(){"}"===Ge.id&&et("Unexpected empty array");let e=[],t=!1;const n=[e],l=ft(!1);for(;!t;){if(Ze(Ge)&&nt(),De(Ge))e.push(Ye.Literal.nud.call(Ge)),nt();else if(Qe&&He(Ge))e.push(Ye[i].nud.call(Ge)),nt();else if(Je&&qe(Ge)){const t=lt(0);e.push(t)}else et(`Unexpected ${Ge.type} in array: ${Ge.value}`);","===Ge.id?nt(","):";"===Ge.id?(nt(";"),e=[],n.push(e)):t=!0}const r=Ge;return nt("}"),ft(l),this.type="ArrayExpression",this.elements=n,this.loc&&(this.loc[1]=r.loc[1]),delete this.value,this}));const yt=(e,t,n)=>Math.min(Math.max(t,e),n);function mt(e,t){return t?String(e+1):e?"["+e+"]":""}function vt(e){let{r0:t,c0:n,r1:l,c1:r}=e;const{$c0:o,$c1:u,$r0:s,$r1:c}=e,i=null==t,a=null==n;let f=null==l,p=null==r;t=yt(s?0:-1048575,0|t,1048575),n=yt(o?0:-16383,0|n,16383),!i&&f&&!a&&p?(l=t,f=!1,r=n,p=!1):(l=yt(c?0:-1048575,0|l,1048575),r=yt(u?0:-16383,0|r,16383));if(0===t&&l>=1048575&&!a&&!p||i&&f){const e=mt(n,o),t=mt(r,u);return"C"+(e===t?e:e+":C"+t)}if(0===n&&r>=16383&&!i&&!f||a&&p){const e=mt(t,s),n=mt(l,c);return"R"+(e===n?e:e+":R"+n)}const h=mt(t,s),g=mt(l,c),x=mt(n,o),d=mt(r,u);return i||f||a||p?(i?"":"R"+h)+(a?"":"C"+x)+":"+(f?"":"R"+g)+(p?"":"C"+d):h!==g||x!==d?"R"+h+"C"+x+":R"+g+"C"+d:"R"+h+"C"+x}function Et(e){let t=null,n=null,l=null,r=null;const o=/^R(?:\[([+-]?\d+)\]|(\d+))?/.exec(e);o&&(o[1]?(t=parseInt(o[1],10),l=!1):o[2]?(t=parseInt(o[2],10)-1,l=!0):(t=0,l=!1),e=e.slice(o[0].length));const u=/^C(?:\[([+-]?\d+)\]|(\d+))?/.exec(e);return u&&(u[1]?(n=parseInt(u[1],10),r=!1):u[2]?(n=parseInt(u[2],10)-1,r=!0):(n=0,r=!1),e=e.slice(u[0].length)),!o&&!u||e.length?null:[t,n,l,r]}function Rt(e){let t=null;const[n,l]=e.split(":",2),r=Et(n);if(r){const[e,n,o,u]=r;if(!l)return null!=e&&null==n?{r0:e,c0:null,r1:e,c1:null,$r0:o,$c0:!1,$r1:o,$c1:!1}:null==e&&null!=n?{r0:null,c0:n,r1:null,c1:n,$r0:!1,$c0:u,$r1:!1,$c1:u}:{r0:e||0,c0:n||0,r1:e||0,c1:n||0,$r0:o||!1,$c0:u||!1,$r1:o||!1,$c1:u||!1};{const r=Et(l);if(!r)return null;{t={};const[l,s,c,i]=r;null!=e&&null!=l?(t.r0=o===c?Math.min(e,l):e,t.$r0=o,t.r1=o===c?Math.max(e,l):l,t.$r1=c):null!=e&&null==l?(t.r0=e,t.$r0=o,t.r1=null,t.$r1=o):null==e&&null!=l?(t.r0=l,t.$r0=c,t.r1=null,t.$r1=c):null==e&&null==l&&(t.r0=null,t.$r0=!1,t.r1=null,t.$r1=!1),null!=n&&null!=s?(t.c0=u===i?Math.min(n,s):n,t.$c0=u,t.c1=u===i?Math.max(n,s):s,t.$c1=i):null!=n&&null==s?(t.c0=n,t.$c0=u,t.c1=null,t.$c1=u):null==n&&null!=s?(t.c0=s,t.$c0=i,t.c1=null,t.$c1=i):null==n&&null==s&&(t.c0=null,t.$c0=!1,t.c1=null,t.$c1=!1)}}}return t}function wt(e){let{allowNamed:t=!0,allowTernary:n=!1,xlsx:l=!1}=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};const r=b(e,{allowNamed:t,allowTernary:n,xlsx:l,r1c1:!0});if(r&&(r.r0||r.name)){const e=r.r1?Rt(r.r0+":"+r.r1):Rt(r.r0);return r.name||e?(r.range=e,delete r.r0,delete r.r1,r):null}return null}function Nt(e){let{xlsx:t=!1}=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};const n=t?T(e):C(e);return n+(e.name?e.name:vt(e.range))}const bt=(e,t,n)=>null==t?null:e?t:t-n,At={withLocation:!1,mergeRefs:!1,allowTernary:!0,r1c1:!1};function Ct(e,t,n,l){let r=!(arguments.length>4&&void 0!==arguments[4])||arguments[4],o=e;if(null!=o&&!t){if(o=n+e,o<0){if(!r)return NaN;o=l+o+1}if(o>l){if(!r)return NaN;o-=l+1}}return o}const Tt={wrapEdges:!0,mergeRefs:!0,allowTernary:!0,xlsx:!1};const It=Object.freeze({OPERATOR:e,BOOLEAN:"bool",ERROR:t,NUMBER:"number",FUNCTION:"func",NEWLINE:"newline",WHITESPACE:"whitespace",STRING:"string",CONTEXT:"context",CONTEXT_QUOTE:"context_quote",REF_RANGE:"range",REF_BEAM:n,REF_TERNARY:l,REF_NAMED:r,REF_STRUCT:o,FX_PREFIX:"fx_prefix",UNKNOWN:u}),Lt=Object.freeze({UNARY:s,BINARY:c,REFERENCE:i,LITERAL:"Literal",ERROR:"ErrorLiteral",CALL:a,ARRAY:"ArrayExpression",IDENTIFIER:"Identifier"});exports.MAX_COLS=16383,exports.MAX_ROWS=1048575,exports.addA1RangeBounds=j,exports.addTokenMeta=function(e){let{sheetName:t="",workbookName:r=""}=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};const s=[];let c=null;const i=pt(),a=[],f=()=>s.length+(c?1:0);return e.forEach(((e,p)=>{if(e.index=p,e.depth=f(),"("===e.value)s.push(e),e.depth=f();else if(")"===e.value){const t=s.pop();if(t){const n=i();e.groupId=n,e.depth=t.depth,t.groupId=n}else e.error=!0}else if("{"===e.value)c?e.error=!0:(c=e,e.depth=f());else if("}"===e.value){if(c){const t=i();e.groupId=t,e.depth=c.depth,c.groupId=t}else e.error=!0;c=null}else if("range"===e.type||e.type===n||e.type===l||e.type===o){const n=e.type===o?Y(e.value,{allowTernary:!0,xlsx:!0}):z(e.value,{allowTernary:!0,xlsx:!0});if(n&&(n.range||n.columns)){n.source=e.value,$t(n,t,r);const l=a.find((e=>dt(e,n)));l?e.groupId=l.groupId:(n.groupId=i(),e.groupId=n.groupId,a.push(n))}}else e.type===u&&(e.error=!0)})),e},exports.fixRanges=function e(t){let n=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{addBounds:!1};if("string"==typeof t)return e(Se(t,n),n).map((e=>e.value)).join("");if(!Array.isArray(t))throw new Error("fixRanges expects an array of tokens");const{addBounds:l,r1c1:r,xlsx:u}=n;if(r)throw new Error("fixRanges does not have an R1C1 mode");let s=0;return t.map((e=>{const t={...e};e.loc&&(t.loc=[...e.loc]);let n=0;if(t.type===o){const e=Q(Y(t.value,{xlsx:u}),{xlsx:u});n=e.length-t.value.length,t.value=e}else if(Me(t)){const e=z(t.value,{xlsx:u,allowTernary:!0}),r=e.range;l&&j(r);const o=D(e,{xlsx:u});n=o.length-t.value.length,t.value=o}return s||n?(t.loc&&(t.loc[0]+=s),s+=n,t.loc&&(t.loc[1]+=s)):s+=n,t}))},exports.fromCol=_,exports.isError=je,exports.isFunction=qe,exports.isFxPrefix=Be,exports.isLiteral=De,exports.isOperator=Pe,exports.isRange=Me,exports.isReference=ze,exports.isWhitespace=Ze,exports.mergeRefTokens=Le,exports.nodeTypes=Lt,exports.parse=function(e,t){if("string"==typeof e)Ke=Se(e,{withLocation:!1,...t,mergeRefs:!0});else{if(!Array.isArray(e))throw new Error("Parse requires a string or array of tokens.");Ke=e}for(Qe=t?.permitArrayRanges,Je=t?.permitArrayCalls,Ve=0;Ze(Ke[Ve])||Be(Ke[Ve]);)Ve++;nt(),ft(!0);const n=lt(0);return nt(Xe),n},exports.parseA1Ref=z,exports.parseR1C1Ref=wt,exports.parseStructRef=Y,exports.stringifyA1Ref=D,exports.stringifyR1C1Ref=Nt,exports.stringifyStructRef=Q,exports.toCol=U,exports.tokenTypes=It,exports.tokenize=Se,exports.translateToA1=function(e,n){let l=arguments.length>2&&void 0!==arguments[2]?arguments[2]:Tt;const r=M(n),o="string"==typeof e,u={...Tt,...l},s=o?Se(e,{withLocation:!1,mergeRefs:u.mergeRefs,xlsx:u.xlsx,allowTernary:u.allowTernary,r1c1:!0}):e;let c=0;const i={xlsx:u.xlsx,allowTernary:u.allowTernary};return s.forEach((e=>{if(Me(e)){const n=e.value,l=wt(n,i),o=l.range,s={},a=Ct(o.r0,o.$r0,r.top,1048575,u.wrapEdges),f=Ct(o.r1,o.$r1,r.top,1048575,u.wrapEdges);a>f?(s.top=f,s.$top=o.$r1,s.bottom=a,s.$bottom=o.$r0):(s.top=a,s.$top=o.$r0,s.bottom=f,s.$bottom=o.$r1);const p=Ct(o.c0,o.$c0,r.left,16383,u.wrapEdges),h=Ct(o.c1,o.$c1,r.left,16383,u.wrapEdges);p>h?(s.left=h,s.$left=o.$c1,s.right=p,s.$right=o.$c0):(s.left=p,s.$left=o.$c0,s.right=h,s.$right=o.$c1),isNaN(a)||isNaN(f)||isNaN(p)||isNaN(h)?(e.type=t,e.value="#REF!",delete e.groupId):(l.range=s,e.value=D(l,i)),e.loc&&(e.loc[0]+=c,c+=e.value.length-n.length,e.loc[1]+=c)}else c&&e.loc&&(e.loc[0]+=c,e.loc[1]+=c)})),o?s.map((e=>e.value)).join(""):s},exports.translateToR1C1=function(e,t){let{xlsx:n=!1,allowTernary:l=!0}=arguments.length>2&&void 0!==arguments[2]?arguments[2]:{};const{top:r,left:o}=M(t),u="string"==typeof e,s=u?Se(e,{...At,xlsx:n,allowTernary:l}):e;let c=0;const i={xlsx:n,allowTernary:l};return s.forEach((e=>{if(Me(e)){const t=e.value,n=z(t,i),l=n.range,u={};u.r0=bt(l.$top,l.top,r),u.r1=bt(l.$bottom,l.bottom,r),u.c0=bt(l.$left,l.left,o),u.c1=bt(l.$right,l.right,o),u.$r0=l.$top,u.$r1=l.$bottom,u.$c0=l.$left,u.$c1=l.$right,n.range=u,e.value=Nt(n,i),e.loc&&(e.loc[0]+=c,c+=e.value.length-t.length,e.loc[1]+=c)}else c&&e.loc&&(e.loc[0]+=c,e.loc[1]+=c)})),u?s.map((e=>e.value)).join(""):s};
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiZnguanMiLCJzb3VyY2VzIjpbXSwic291cmNlc0NvbnRlbnQiOltdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiIn0=
