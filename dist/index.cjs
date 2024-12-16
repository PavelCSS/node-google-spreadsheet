"use strict";const compact_js=require("lodash/compact.js"),each_js=require("lodash/each.js"),filter_js=require("lodash/filter.js"),find_js=require("lodash/find.js"),flatten_js=require("lodash/flatten.js"),get_js=require("lodash/get.js"),groupBy_js=require("lodash/groupBy.js"),isArray_js=require("lodash/isArray.js"),isBoolean_js=require("lodash/isBoolean.js"),isEqual_js=require("lodash/isEqual.js"),isFinite_js=require("lodash/isFinite.js"),isInteger_js=require("lodash/isInteger.js"),isNil_js=require("lodash/isNil.js");require("lodash/isNumber.js");const isObject_js=require("lodash/isObject.js"),isString_js=require("lodash/isString.js"),keyBy_js=require("lodash/keyBy.js"),keys_js=require("lodash/keys.js"),map_js=require("lodash/map.js"),omit_js=require("lodash/omit.js"),pickBy_js=require("lodash/pickBy.js"),set_js=require("lodash/set.js"),some_js=require("lodash/some.js"),sortBy_js=require("lodash/sortBy.js"),times_js=require("lodash/times.js"),unset_js=require("lodash/unset.js"),values_js=require("lodash/values.js"),merge_js=require("lodash/merge.js"),p=require("axios");function _interopDefaultCompat(r){return r&&typeof r=="object"&&"default"in r?r.default:r}const compact_js__default=_interopDefaultCompat(compact_js),each_js__default=_interopDefaultCompat(each_js),filter_js__default=_interopDefaultCompat(filter_js),find_js__default=_interopDefaultCompat(find_js),flatten_js__default=_interopDefaultCompat(flatten_js),get_js__default=_interopDefaultCompat(get_js),groupBy_js__default=_interopDefaultCompat(groupBy_js),isArray_js__default=_interopDefaultCompat(isArray_js),isBoolean_js__default=_interopDefaultCompat(isBoolean_js),isEqual_js__default=_interopDefaultCompat(isEqual_js),isFinite_js__default=_interopDefaultCompat(isFinite_js),isInteger_js__default=_interopDefaultCompat(isInteger_js),isNil_js__default=_interopDefaultCompat(isNil_js),isObject_js__default=_interopDefaultCompat(isObject_js),isString_js__default=_interopDefaultCompat(isString_js),keyBy_js__default=_interopDefaultCompat(keyBy_js),keys_js__default=_interopDefaultCompat(keys_js),map_js__default=_interopDefaultCompat(map_js),omit_js__default=_interopDefaultCompat(omit_js),pickBy_js__default=_interopDefaultCompat(pickBy_js),set_js__default=_interopDefaultCompat(set_js),some_js__default=_interopDefaultCompat(some_js),sortBy_js__default=_interopDefaultCompat(sortBy_js),times_js__default=_interopDefaultCompat(times_js),unset_js__default=_interopDefaultCompat(unset_js),values_js__default=_interopDefaultCompat(values_js),merge_js__default=_interopDefaultCompat(merge_js),p__default=_interopDefaultCompat(p);function getFieldMask(r){let e="";const t=Object.keys(r).filter(s=>s!=="gridProperties").join(",");return r.gridProperties&&(e=Object.keys(r.gridProperties).map(s=>`gridProperties.${s}`).join(","),e.length&&t.length&&(e=`${e},`)),e+t}function columnToLetter(r){let e,t="",s=r;for(;s>0;)e=(s-1)%26,t=String.fromCharCode(e+65)+t,s=(s-e-1)/26;return t}function letterToColumn(r){let e=0;const{length:t}=r;for(let s=0;s<t;s++)e+=(r.charCodeAt(s)-64)*26**(t-s-1);return e}function axiosParamsSerializer(r){let e="";return Object.keys(r).forEach(t=>{const s=typeof r[t]=="object",a=s&&r[t].length>=0;if(s||(e+=`${t}=${encodeURIComponent(r[t])}&`),s&&a)for(const n of r[t])e+=`${t}=${encodeURIComponent(n)}&`}),e&&e.slice(0,-1)}function checkForDuplicateHeaders(r){const e=groupBy_js__default(r);each_js__default(e,(t,s)=>{if(s&&t.length>1)throw new Error(`Duplicate header detected: "${s}". Please make sure all non-empty headers are unique`)})}const parseRangeA1=r=>{const e=r.match(/([A-Z]+)?(\d+)?:?([A-Z]+)?(\d+)?$/)||[];if(!e)throw new Error(`Range address "${r}" not valid`);const[t="",s,a,n,o]=e,i=s?letterToColumn(s)-1:0,h=n?letterToColumn(n):void 0,d=a?parseInt(a)-1:0,l=o?parseInt(o):void 0;return{rangeA1:t,startColumnIndex:i,endColumnIndex:h,startRowIndex:d,endRowIndex:l,startColumnA1:s,endColumnA1:n,startRowA1:a,endRowA1:o}},toA1Range=({startColumnIndex:r,endColumnIndex:e,startRowIndex:t,endRowIndex:s})=>{let a="",n="";return Number.isInteger(r)&&(a+=columnToLetter(r+1)),Number.isInteger(t)&&(a+=`${t+1}`),Number.isInteger(e)&&(n+=columnToLetter(e)),Number.isInteger(s)&&(n+=`${s}`),a===n||a&&!n?a:[a,n].join(":")};class GoogleSpreadsheetRow{constructor(e,t,s){this._worksheet=e,this._rowNumber=t,this._rawData=s,this._deleted=!1}get deleted(){return this._deleted}get rowNumber(){return this._rowNumber}_updateRowNumber(e){this._rowNumber=e}get a1Range(){return[this._worksheet.a1SheetName,"!",`A${this._rowNumber}`,":",`${columnToLetter(this._worksheet.headerValues.length)}${this._rowNumber}`].join("")}get(e){const t=this._worksheet.headerValues.indexOf(e);return this._rawData[t]}set(e,t){const s=this._worksheet.headerValues.indexOf(e);this._rawData[s]=t}assign(e){for(const t in e)this.set(t,e[t])}toObject(){const e={};for(let t=0;t<this._worksheet.headerValues.length;t++){const s=this._worksheet.headerValues[t];s&&(e[s]=this._rawData[t])}return e}async save(e){if(this._deleted)throw new Error("This row has been deleted - call getRows again before making updates.");const t=await this._worksheet._spreadsheet.sheetsApi.request({method:"put",url:`/values/${encodeURIComponent(this.a1Range)}`,params:{valueInputOption:e?.raw?"RAW":"USER_ENTERED",includeValuesInResponse:!0},data:{range:this.a1Range,majorDimension:"ROWS",values:[this._rawData]}});this._rawData=t.data.updatedData.values[0]}async delete(){if(this._deleted)throw new Error("This row has been deleted - call getRows again before making updates.");const e=await this._worksheet._makeSingleUpdateRequest("deleteRange",{range:{sheetId:this._worksheet.sheetId,startRowIndex:this._rowNumber-1,endRowIndex:this._rowNumber},shiftDimension:"ROWS"});return this._deleted=!0,this._worksheet._shiftRowCache(this.rowNumber),e}_clearRowData(){for(let e=0;e<this._rawData.length;e++)this._rawData[e]=""}}class GoogleSpreadsheetCellErrorValue{constructor(e){this.type=e.type,this.message=e.message}}class GoogleSpreadsheetCell{constructor(e,t,s,a){this._sheet=e,this._rowIndex=t,this._columnIndex=s,this._draftData={},this._updateRawData(a),this._rawData=a}_updateRawData(e){this._rawData=e,this._draftData={},this._rawData?.effectiveValue&&"errorValue"in this._rawData.effectiveValue?this._error=new GoogleSpreadsheetCellErrorValue(this._rawData.effectiveValue.errorValue):this._error=void 0}get rowIndex(){return this._rowIndex}get columnIndex(){return this._columnIndex}get a1Column(){return columnToLetter(this._columnIndex+1)}get a1Row(){return this._rowIndex+1}get a1Address(){return`${this.a1Column}${this.a1Row}`}get value(){if(this._draftData.value!==void 0)throw new Error("Value has been changed");return this._error?this._error:this._rawData?.effectiveValue?values_js__default(this._rawData.effectiveValue)[0]:null}set value(e){if(e instanceof GoogleSpreadsheetCellErrorValue)throw new Error("You can't manually set a value to an error");if(isBoolean_js__default(e))this._draftData.valueType="boolValue";else if(isString_js__default(e))e.substring(0,1)==="="?this._draftData.valueType="formulaValue":this._draftData.valueType="stringValue";else if(isFinite_js__default(e))this._draftData.valueType="numberValue";else if(isNil_js__default(e))this._draftData.valueType="stringValue",e="";else throw new Error("Set value to boolean, string, or number");this._draftData.value=e}get valueType(){return this._error?"errorValue":this._rawData?.effectiveValue?keys_js__default(this._rawData.effectiveValue)[0]:null}get formattedValue(){return this._rawData?.formattedValue||null}get formula(){return get_js__default(this._rawData,"userEnteredValue.formulaValue",null)}set formula(e){if(!e)throw new Error("To clear a formula, set `cell.value = null`");if(e.substring(0,1)!=="=")throw new Error('formula must begin with "="');this.value=e}get formulaError(){return this._error}get errorValue(){return this._error}get numberValue(){if(this.valueType==="numberValue")return this.value}set numberValue(e){this.value=e}get boolValue(){if(this.valueType==="boolValue")return this.value}set boolValue(e){this.value=e}get stringValue(){if(this.valueType==="stringValue")return this.value}set stringValue(e){if(e?.startsWith("="))throw new Error("Use cell.formula to set formula values");this.value=e}get hyperlink(){if(this._draftData.value)throw new Error("Save cell to be able to read hyperlink");return this._rawData?.hyperlink}get note(){return this._draftData.note!==void 0?this._draftData.note:this._rawData?.note}set note(e){if((e==null||e===!1)&&(e=""),!isString_js__default(e))throw new Error("Note must be a string");e===this._rawData?.note?delete this._draftData.note:this._draftData.note=e}get userEnteredFormat(){return Object.freeze(this._rawData?.userEnteredFormat)}get effectiveFormat(){return Object.freeze(this._rawData?.effectiveFormat)}_getFormatParam(e){if(get_js__default(this._draftData,`userEnteredFormat.${e}`))throw new Error("User format is unsaved - save the cell to be able to read it again");return Object.freeze(this._rawData.userEnteredFormat[e])}_setFormatParam(e,t){isEqual_js__default(t,get_js__default(this._rawData,`userEnteredFormat.${e}`))?unset_js__default(this._draftData,`userEnteredFormat.${e}`):(set_js__default(this._draftData,`userEnteredFormat.${e}`,t),this._draftData.clearFormat=!1)}get numberFormat(){return this._getFormatParam("numberFormat")}get backgroundColor(){return this._getFormatParam("backgroundColor")}get backgroundColorStyle(){return this._getFormatParam("backgroundColorStyle")}get borders(){return this._getFormatParam("borders")}get padding(){return this._getFormatParam("padding")}get horizontalAlignment(){return this._getFormatParam("horizontalAlignment")}get verticalAlignment(){return this._getFormatParam("verticalAlignment")}get wrapStrategy(){return this._getFormatParam("wrapStrategy")}get textDirection(){return this._getFormatParam("textDirection")}get textFormat(){return this._getFormatParam("textFormat")}get hyperlinkDisplayType(){return this._getFormatParam("hyperlinkDisplayType")}get textRotation(){return this._getFormatParam("textRotation")}set numberFormat(e){this._setFormatParam("numberFormat",e)}set backgroundColor(e){this._setFormatParam("backgroundColor",e)}set backgroundColorStyle(e){this._setFormatParam("backgroundColorStyle",e)}set borders(e){this._setFormatParam("borders",e)}set padding(e){this._setFormatParam("padding",e)}set horizontalAlignment(e){this._setFormatParam("horizontalAlignment",e)}set verticalAlignment(e){this._setFormatParam("verticalAlignment",e)}set wrapStrategy(e){this._setFormatParam("wrapStrategy",e)}set textDirection(e){this._setFormatParam("textDirection",e)}set textFormat(e){this._setFormatParam("textFormat",e)}set hyperlinkDisplayType(e){this._setFormatParam("hyperlinkDisplayType",e)}set textRotation(e){this._setFormatParam("textRotation",e)}clearAllFormatting(){this._draftData.clearFormat=!0,delete this._draftData.userEnteredFormat}get _isDirty(){return!!(this._draftData.note!==void 0||keys_js__default(this._draftData.userEnteredFormat).length||this._draftData.clearFormat||this._draftData.value!==void 0)}discardUnsavedChanges(){this._draftData={}}async save(){await this._sheet.saveCells([this])}_getUpdateRequest(){const e=this._draftData.value!==void 0,t=this._draftData.note!==void 0,s=!!keys_js__default(this._draftData.userEnteredFormat||{}).length,a=this._draftData.clearFormat;if(!some_js__default([e,t,s,a]))return null;const n={...this._rawData?.userEnteredFormat,...this._draftData.userEnteredFormat};return get_js__default(this._draftData,"userEnteredFormat.backgroundColor")&&delete n.backgroundColorStyle,{updateCells:{rows:[{values:[{...e&&{userEnteredValue:{[this._draftData.valueType]:this._draftData.value}},...t&&{note:this._draftData.note},...s&&{userEnteredFormat:n},...a&&{userEnteredFormat:{}}}]}],fields:keys_js__default(pickBy_js__default({userEnteredValue:e,note:t,userEnteredFormat:s||a})).join(","),start:{sheetId:this._sheet.sheetId,rowIndex:this.rowIndex,columnIndex:this.columnIndex}}}}}const C=r=>r.find(e=>e).map((e,t)=>r.map(s=>s[t]));class GoogleSpreadsheetWorksheet{constructor(e,t,s){this._spreadsheet=e,this._headerRowIndex=1,this._rawProperties=null,this._cells=[],this._rowMetadata=[],this._columnMetadata=[],this._rowCache=[],this._headerRowIndex=1,this._rawProperties=t,this._cells=[],this._rowMetadata=[],this._columnMetadata=[],s&&this._fillCellData(s)}get headerValues(){if(!this._headerValues)throw new Error("Header values are not yet loaded");return this._headerValues}updateRawData(e,t){this._rawProperties=e,this._fillCellData(t)}async _makeSingleUpdateRequest(e,t){return this._spreadsheet._makeSingleUpdateRequest(e,{...t})}_ensureInfoLoaded(){if(!this._rawProperties)throw new Error("You must call `doc.loadInfo()` again before accessing this property")}resetLocalCache(e){e||(this._rawProperties=null),this._headerValues=void 0,this._headerRowIndex=1,this._cells=[]}_fillCellData(e){each_js__default(e,t=>{const s=t.startRow||0,a=t.startColumn||0,n=t.rowMetadata.length,o=t.columnMetadata.length;for(let i=0;i<n;i++){const h=s+i;for(let d=0;d<o;d++){const l=a+d;this._cells[h]||(this._cells[h]=[]);const _=get_js__default(t,`rowData[${i}].values[${d}]`);this._cells[h][l]?this._cells[h][l]._updateRawData(_):this._cells[h][l]=new GoogleSpreadsheetCell(this,h,l,_)}}for(let i=0;i<t.rowMetadata.length;i++)this._rowMetadata[s+i]=t.rowMetadata[i];for(let i=0;i<t.columnMetadata.length;i++)this._columnMetadata[a+i]=t.columnMetadata[i]})}_addSheetIdToRange(e){if(e.sheetId&&e.sheetId!==this.sheetId)throw new Error("Leave sheet ID blank or set to matching ID of this sheet");return{...e,sheetId:this.sheetId}}_getProp(e){return this._ensureInfoLoaded(),this._rawProperties[e]}_setProp(e,t){throw new Error("Do not update directly - use `updateProperties()`")}get sheetId(){return this._getProp("sheetId")}get title(){return this._getProp("title")}get index(){return this._getProp("index")}get sheetType(){return this._getProp("sheetType")}get gridProperties(){return this._getProp("gridProperties")}get hidden(){return this._getProp("hidden")}get tabColor(){return this._getProp("tabColor")}get rightToLeft(){return this._getProp("rightToLeft")}get _headerRange(){return`A${this._headerRowIndex}:${this.lastColumnLetter}${this._headerRowIndex}`}set sheetId(e){this._setProp("sheetId",e)}set title(e){this._setProp("title",e)}set index(e){this._setProp("index",e)}set sheetType(e){this._setProp("sheetType",e)}set gridProperties(e){this._setProp("gridProperties",e)}set hidden(e){this._setProp("hidden",e)}set tabColor(e){this._setProp("tabColor",e)}set rightToLeft(e){this._setProp("rightToLeft",e)}get rowCount(){return this._ensureInfoLoaded(),this.gridProperties.rowCount}get columnCount(){return this._ensureInfoLoaded(),this.gridProperties.columnCount}get a1SheetName(){return`'${this.title.replace(/'/g,"''")}'`}get encodedA1SheetName(){return encodeURIComponent(this.a1SheetName)}get lastColumnLetter(){return this.columnCount?columnToLetter(this.columnCount):""}get cellStats(){let e=flatten_js__default(this._cells);return e=compact_js__default(e),{nonEmpty:filter_js__default(e,t=>t.value).length,loaded:e.length,total:this.rowCount*this.columnCount}}getCellByA1(e){const t=e.match(/([A-Z]+)([0-9]+)/);if(!t)throw new Error(`Cell address "${e}" not valid`);const s=letterToColumn(t[1]),a=parseInt(t[2]);return this.getCell(a-1,s-1)}getCell(e,t){if(e<0||t<0)throw new Error("Min coordinate is 0, 0");if(e>=this.rowCount||t>=this.columnCount)throw new Error(`Out of bounds, sheet is ${this.rowCount} by ${this.columnCount}`);if(!get_js__default(this._cells,`[${e}][${t}]`))throw new Error("This cell has not been loaded yet");return this._cells[e][t]}getCellsByA1Range(e,t=!1){const{startRowIndex:s,startColumnIndex:a,endRowIndex:n,endColumnIndex:o}=parseRangeA1(e);return this.getCells({startRowIndex:s,startColumnIndex:a,endRowIndex:n,endColumnIndex:o},t)}getCells({startRowIndex:e=0,startColumnIndex:t=0,endRowIndex:s,endColumnIndex:a},n=!1){let o=this._cells.slice(e,s).map(i=>i.slice(t,a));if(n&&(o=C(o)),Number.isInteger(n?a:s)||(o=o.filter(i=>!!i.find(h=>!!h.formattedValue))),!Number.isInteger(n?s:a)){const i=o.reduce((h,d)=>{const l=d.findIndex(w=>!!w.formattedValue),_=d.length-l;return h>_?h:_},0);o=o.map(h=>h.slice(0,i))}return o}async loadCells(e){if(!e)return this._spreadsheet.loadCells(this.a1SheetName);const t=isArray_js__default(e)?e:[e],s=map_js__default(t,a=>{if(isString_js__default(a))return a.startsWith(this.a1SheetName)?a:`${this.a1SheetName}!${a}`;if(isObject_js__default(a)){const n=a;if(n.sheetId&&n.sheetId!==this.sheetId)throw new Error("Leave sheet ID blank or set to matching ID of this sheet");return{sheetId:this.sheetId,...a}}throw new Error("Each filter must be a A1 range string or gridrange object")});return this._spreadsheet.loadCells(s)}async saveUpdatedCells(){const e=filter_js__default(flatten_js__default(this._cells),{_isDirty:!0});e.length&&await this.saveCells(e)}async saveCells(e){const t=map_js__default(e,a=>a._getUpdateRequest()),s=map_js__default(e,a=>`${this.a1SheetName}!${a.a1Address}`);if(!compact_js__default(t).length)throw new Error("At least one cell must have something to update");await this._spreadsheet._makeBatchUpdateRequest(t,s)}async _ensureHeaderRowLoaded(){this._headerValues||await this.loadHeaderRow()}async loadHeaderRow(e){e!==void 0&&(this._headerRowIndex=e);const t=await this.getCellsInRange(this._headerRange);this._processHeaderRow(t)}_processHeaderRow(e){if(!e)throw new Error("No values in the header row - fill the first row with header values before trying to interact with rows");if(this._headerValues=map_js__default(e[0],t=>t?.trim()),!compact_js__default(this.headerValues).length)throw new Error("All your header cells are blank - fill the first row with header values before trying to interact with rows");checkForDuplicateHeaders(this.headerValues)}async setHeaderRow(e,t){if(!e)return;if(e.length>this.columnCount)throw new Error(`Sheet is not large enough to fit ${e.length} columns. Resize the sheet first.`);const s=map_js__default(e,n=>n?.trim());if(checkForDuplicateHeaders(s),!compact_js__default(s).length)throw new Error("All your header cells are blank -");t&&(this._headerRowIndex=t);const a=await this._spreadsheet.sheetsApi.request({method:"put",url:`/values/${this.encodedA1SheetName}!${this._headerRowIndex}:${this._headerRowIndex}`,params:{valueInputOption:"USER_ENTERED",includeValuesInResponse:!0},data:{range:`${this.a1SheetName}!${this._headerRowIndex}:${this._headerRowIndex}`,majorDimension:"ROWS",values:[[...s,...times_js__default(this.columnCount-s.length,()=>"")]]}});this._headerValues=a.data.updatedData.values[0]}async addRows(e,t={}){if(this.title.includes(":"))throw new Error('Please remove the ":" from your sheet title. There is a bug with the google API which breaks appending rows if any colons are in the sheet title.');if(!isArray_js__default(e))throw new Error("You must pass in an array of row values to append");await this._ensureHeaderRowLoaded();const s=[];each_js__default(e,i=>{let h;if(isArray_js__default(i))h=i;else if(isObject_js__default(i)){h=[];for(let d=0;d<this.headerValues.length;d++){const l=this.headerValues[d];h[d]=i[l]}}else throw new Error("Each row must be an object or an array");s.push(h)});const a=await this._spreadsheet.sheetsApi.request({method:"post",url:`/values/${this.encodedA1SheetName}!A${this._headerRowIndex}:append`,params:{valueInputOption:t.raw?"RAW":"USER_ENTERED",insertDataOption:t.insert?"INSERT_ROWS":"OVERWRITE",includeValuesInResponse:!0},data:{values:s}}),{updatedRange:n}=a.data.updates;let o=n.match(/![A-Z]+([0-9]+):?/)[1];return o=parseInt(o),this._ensureInfoLoaded(),t.insert?this._rawProperties.gridProperties.rowCount+=e.length:o+e.length>this.rowCount&&(this._rawProperties.gridProperties.rowCount=o+e.length-1),map_js__default(a.data.updates.updatedData.values,i=>new GoogleSpreadsheetRow(this,o++,i))}async addRow(e,t){return(await this.addRows([e],t))[0]}async getRows(e){const t=e?.offset||0,s=e?.limit||this.rowCount-1,a=1+this._headerRowIndex+t,n=a+s-1;let o;if(this._headerValues){const d=columnToLetter(this.headerValues.length);o=await this.getCellsInRange(`A${a}:${d}${n}`)}else{const d=await this.batchGetCellsInRange([this._headerRange,`A${a}:${this.lastColumnLetter}${n}`]);this._processHeaderRow(d[0]),o=d[1]}if(!o)return[];const i=[];let h=a;for(let d=0;d<o.length;d++){const l=new GoogleSpreadsheetRow(this,h++,o[d]);this._rowCache[l.rowNumber]=l,i.push(l)}return i}async deleteRows(e){let t=Array.isArray(e)?e:[e];if(!t.length)return;t=t.filter(a=>{if(a.deleted)throw new Error("This row has been deleted - call getRows again before making updates.");return!a.deleted});const s=t.map(a=>({deleteRange:{range:{sheetId:this.sheetId,startRowIndex:a.rowNumber-1,endRowIndex:a.rowNumber},shiftDimension:"ROWS"}}));await this._spreadsheet._makeBatchUpdateRequest(s),t.forEach(a=>this._shiftRowCache(a.rowNumber))}async setValues(e,t,s="ROWS"){const a=`${this.encodedA1SheetName}!${isString_js__default(e)?e:toA1Range(e)}`;await this._spreadsheet.sheetsApi.request({method:"put",url:`/values/${a}`,params:{valueInputOption:"RAW",includeValuesInResponse:!0},data:{range:a,values:t,majorDimension:s}})}_shiftRowCache(e){delete this._rowCache[e],this._rowCache.forEach(t=>{t.rowNumber>e&&t._updateRowNumber(t.rowNumber-1)})}async clearRows(e){const t=e?.start||this._headerRowIndex+1,s=e?.end||this.rowCount;await this._spreadsheet.sheetsApi.post(`/values/${this.encodedA1SheetName}!${t}:${s}:clear`),this._rowCache.forEach(a=>{a.rowNumber>=t&&a.rowNumber<=s&&a._clearRowData()})}async updateProperties(e){return this._makeSingleUpdateRequest("updateSheetProperties",{properties:{sheetId:this.sheetId,...e},fields:getFieldMask(e)})}async updateGridProperties(e){return this.updateProperties({gridProperties:e})}async resize(e){return this.updateGridProperties(e)}async updateDimensionProperties(e,t,s){return this._makeSingleUpdateRequest("updateDimensionProperties",{range:{sheetId:this.sheetId,dimension:e,...s},properties:t,fields:getFieldMask(t)})}async getCellsInRange(e,t){return(await this._spreadsheet.sheetsApi.get(`/values/${this.encodedA1SheetName}!${e}`,{params:t})).data.values}async batchGetCellsInRange(e,t){const s=e.map(a=>`ranges=${this.encodedA1SheetName}!${a}`).join("&");return(await this._spreadsheet.sheetsApi.get(`/values:batchGet?${s}`,{params:t})).data.valueRanges.map(a=>a.values)}async updateNamedRange(){}async addNamedRange(e,t){return this._spreadsheet.addNamedRange(e,{...t,sheetId:this.sheetId})}async deleteNamedRange(){}async repeatCell(){}async autoFill(){}async cutPaste(){}async copyPaste(){}async mergeCells(e,t="MERGE_ALL"){await this._makeSingleUpdateRequest("mergeCells",{mergeType:t,range:this._addSheetIdToRange(e)})}async unmergeCells(e){await this._makeSingleUpdateRequest("unmergeCells",{range:this._addSheetIdToRange(e)})}async updateBorders(){}async addFilterView(){}async appendCells(){}async clearBasicFilter(){}async deleteDimension(){}async deleteEmbeddedObject(){}async deleteFilterView(){}async duplicateFilterView(){}async duplicate(e){const t=(await this._makeSingleUpdateRequest("duplicateSheet",{sourceSheetId:this.sheetId,...e?.index!==void 0&&{insertSheetIndex:e.index},...e?.id&&{newSheetId:e.id},...e?.title&&{newSheetName:e.title}})).properties.sheetId;return this._spreadsheet.sheetsById[t]}async findReplace(){}async insertDimension(e,t,s){if(!e)throw new Error("You need to specify a dimension. i.e. COLUMNS|ROWS");if(!isObject_js__default(t))throw new Error("`range` must be an object containing `startIndex` and `endIndex`");if(!isInteger_js__default(t.startIndex)||t.startIndex<0)throw new Error("range.startIndex must be an integer >=0");if(!isInteger_js__default(t.endIndex)||t.endIndex<0)throw new Error("range.endIndex must be an integer >=0");if(t.endIndex<=t.startIndex)throw new Error("range.endIndex must be greater than range.startIndex");if(s===void 0&&(s=t.startIndex>0),s&&t.startIndex===0)throw new Error("Cannot set inheritFromBefore to true if inserting in first row/column");return this._makeSingleUpdateRequest("insertDimension",{range:{sheetId:this.sheetId,dimension:e,startIndex:t.startIndex,endIndex:t.endIndex},inheritFromBefore:s})}async insertRange(){}async moveDimension(){}async updateEmbeddedObjectPosition(){}async pasteData(){}async textToColumns(){}async updateFilterView(){}async deleteRange(){}async appendDimension(){}async addConditionalFormatRule(){}async updateConditionalFormatRule(){}async deleteConditionalFormatRule(){}async sortRange(){}async setDataValidation(e,t){return this._makeSingleUpdateRequest("setDataValidation",{range:{sheetId:this.sheetId,...e},...t&&{rule:t}})}async setBasicFilter(){}async addProtectedRange(){}async updateProtectedRange(){}async deleteProtectedRange(){}async autoResizeDimensions(){}async addChart(){}async updateChartSpec(){}async updateBanding(){}async addBanding(){}async deleteBanding(){}async createDeveloperMetadata(){}async updateDeveloperMetadata(){}async deleteDeveloperMetadata(){}async randomizeRange(){}async addDimensionGroup(){}async deleteDimensionGroup(){}async updateDimensionGroup(){}async trimWhitespace(){}async deleteDuplicates(){}async addSlicer(){}async updateSlicerSpec(){}async delete(){return this._spreadsheet.deleteSheet(this.sheetId)}async copyToSpreadsheet(e){return this._spreadsheet.sheetsApi.post(`/sheets/${this.sheetId}:copyTo`,{destinationSpreadsheetId:e})}async clear(e){const t=e?`!${e}`:"";await this._spreadsheet.sheetsApi.post(`/values/${this.encodedA1SheetName}${t}:clear`),this.resetLocalCache(!0)}async downloadAsCSV(e=!1){return this._spreadsheet._downloadAs("csv",this.sheetId,e)}async downloadAsTSV(e=!1){return this._spreadsheet._downloadAs("tsv",this.sheetId,e)}async downloadAsPDF(e=!1){return this._spreadsheet._downloadAs("pdf",this.sheetId,e)}}var AUTH_MODES=(r=>(r.GOOGLE_AUTH_CLIENT="google_auth",r.RAW_ACCESS_TOKEN="raw_access_token",r.API_KEY="api_key",r))(AUTH_MODES||{});const u="https://sheets.googleapis.com/v4/spreadsheets",A="https://www.googleapis.com/drive/v3/files",c={html:{},zip:{},xlsx:{},ods:{},csv:{singleWorksheet:!0},tsv:{singleWorksheet:!0},pdf:{singleWorksheet:!0}};function m(r){if("getRequestHeaders"in r)return AUTH_MODES.GOOGLE_AUTH_CLIENT;if("token"in r&&r.token)return AUTH_MODES.RAW_ACCESS_TOKEN;if("apiKey"in r&&r.apiKey)return AUTH_MODES.API_KEY;throw new Error("Invalid auth")}async function g(r){if("getRequestHeaders"in r)return{headers:await r.getRequestHeaders()};if("apiKey"in r&&r.apiKey)return{params:{key:r.apiKey}};if("token"in r&&r.token)return{headers:{Authorization:`Bearer ${r.token}`}};throw new Error("Invalid auth")}class GoogleSpreadsheet{constructor(e,t){this._rawProperties=null,this._spreadsheetUrl=null,this._deleted=!1,this.spreadsheetId=e,this.auth=t,this._rawSheets={},this._spreadsheetUrl=null,this.namedRanges={},this.sheetsApi=p__default.create({baseURL:`${u}/${e}`,paramsSerializer:axiosParamsSerializer,maxContentLength:1/0,maxBodyLength:1/0}),this.driveApi=p__default.create({baseURL:`${A}/${e}`,paramsSerializer:axiosParamsSerializer}),this.sheetsApi.interceptors.request.use(this._setAxiosRequestAuth.bind(this)),this.sheetsApi.interceptors.response.use(this._handleAxiosResponse.bind(this),this._handleAxiosErrors.bind(this)),this.driveApi.interceptors.request.use(this._setAxiosRequestAuth.bind(this)),this.driveApi.interceptors.response.use(this._handleAxiosResponse.bind(this),this._handleAxiosErrors.bind(this))}get authMode(){return m(this.auth)}async _setAxiosRequestAuth(e){const t=await g(this.auth);return each_js__default(t.headers,(s,a)=>{e.headers.set(a,s)}),e.params={...e.params,...t.params},e}async _handleAxiosResponse(e){return e}async _handleAxiosErrors(e){const t=e.response?.data;if(t){if(!t.error)throw e;const{code:s,message:a}=t.error;throw e.message=`Google API error - [${s}] ${a}`,e}throw get_js__default(e,"response.status")===403&&"apiKey"in this.auth?new Error("Sheet is private. Use authentication or make public. (see https://github.com/theoephraim/node-google-spreadsheet#a-note-on-authentication for details)"):e}async _makeSingleUpdateRequest(e,t){const s=await this.sheetsApi.post(":batchUpdate",{requests:[{[e]:t}],includeSpreadsheetInResponse:!0});return this._updateNamedRanges(s.data.updatedSpreadsheet.namedRanges),this._updateRawProperties(s.data.updatedSpreadsheet.properties),each_js__default(s.data.updatedSpreadsheet.sheets,a=>this._updateOrCreateSheet(a)),s.data.replies[0][e]}async _makeBatchUpdateRequest(e,t){const s=await this.sheetsApi.post(":batchUpdate",{requests:e,includeSpreadsheetInResponse:!0,...t&&{responseIncludeGridData:!0,...t!=="*"&&{responseRanges:t}}});this._updateNamedRanges(s.data.updatedSpreadsheet.namedRanges),this._updateRawProperties(s.data.updatedSpreadsheet.properties),each_js__default(s.data.updatedSpreadsheet.sheets,a=>this._updateOrCreateSheet(a))}_ensureInfoLoaded(){if(!this._rawProperties)throw new Error("You must call `doc.loadInfo()` before accessing this property")}_updateRawProperties(e){this._rawProperties=e}_updateOrCreateSheet(e){const{properties:t,data:s}=e,{sheetId:a}=t;this._rawSheets[a]?this._rawSheets[a].updateRawData(t,s):this._rawSheets[a]=new GoogleSpreadsheetWorksheet(this,t,s)}_getProp(e){return this._ensureInfoLoaded(),this._rawProperties[e]}get title(){return this._getProp("title")}get locale(){return this._getProp("locale")}get timeZone(){return this._getProp("timeZone")}get autoRecalc(){return this._getProp("autoRecalc")}get defaultFormat(){return this._getProp("defaultFormat")}get spreadsheetTheme(){return this._getProp("spreadsheetTheme")}get iterativeCalculationSettings(){return this._getProp("iterativeCalculationSettings")}async updateProperties(e){await this._makeSingleUpdateRequest("updateSpreadsheetProperties",{properties:e,fields:getFieldMask(e)})}async loadInfo(e=!1){const t=await this.sheetsApi.get("/",{params:{...e&&{includeGridData:!0}}});return this._spreadsheetUrl=t.data.spreadsheetUrl,this._rawProperties=t.data.properties,this._updateNamedRanges(t.data.namedRanges),each_js__default(t.data.sheets,s=>this._updateOrCreateSheet(s)),t}_updateNamedRanges(e){const t=keyBy_js__default(e,"name");this.namedRanges=merge_js__default(this.namedRanges,t)}async updateInfo(e=!1){const t=(await this.loadInfo(e)).data.sheets.map(({properties:s})=>s.sheetId);Object.keys(this._rawSheets).forEach(s=>{t.includes(+s)||delete this._rawSheets[s]})}resetLocalCache(){this._rawProperties=null,this._rawSheets={}}get sheetCount(){return this._ensureInfoLoaded(),values_js__default(this._rawSheets).length}get sheetsById(){return this._ensureInfoLoaded(),this._rawSheets}get sheetsByIndex(){return this._ensureInfoLoaded(),sortBy_js__default(this._rawSheets,"index")}get sheetsByTitle(){return this._ensureInfoLoaded(),keyBy_js__default(this._rawSheets,"title")}async addSheet(e={}){const t=(await this._makeSingleUpdateRequest("addSheet",{properties:omit_js__default(e,"headerValues","headerRowIndex")})).properties.sheetId,s=this.sheetsById[t];return e.headerValues&&await s.setHeaderRow(e.headerValues,e.headerRowIndex),s}async deleteSheet(e){await this._makeSingleUpdateRequest("deleteSheet",{sheetId:e}),delete this._rawSheets[e]}async addNamedRange(e,t){return this._makeSingleUpdateRequest("addNamedRange",{namedRange:{name:e,range:t}})}async deleteNamedRange(e){return this._makeSingleUpdateRequest("deleteNamedRange",{namedRangeId:e})}async loadCells(e){const t=this.authMode===AUTH_MODES.API_KEY,s=isArray_js__default(e)?e:[e],a=map_js__default(s,i=>{if(isString_js__default(i))return t?i:{a1Range:i};if(isObject_js__default(i)){if(t)throw new Error("Only A1 ranges are supported when fetching cells with read-only access (using only an API key)");return{gridRange:i}}throw new Error("Each filter must be an A1 range string or a gridrange object")});let n;this.authMode===AUTH_MODES.API_KEY?n=await this.sheetsApi.get("/",{params:{includeGridData:!0,ranges:a}}):n=await this.sheetsApi.post(":getByDataFilter",{includeGridData:!0,dataFilters:a});const{sheets:o}=n.data;each_js__default(o,i=>{this._updateOrCreateSheet(i)})}async _downloadAs(e,t,s){if(!c[e])throw new Error(`unsupported export fileType - ${e}`);if(c[e].singleWorksheet){if(t===void 0)throw new Error(`Must specify worksheetId when exporting as ${e}`)}else if(t)throw new Error(`Cannot specify worksheetId when exporting as ${e}`);if(e==="html"&&(e="zip"),!this._spreadsheetUrl)throw new Error("Cannot export sheet that is not fully loaded");const a=this._spreadsheetUrl.replace("/edit","/export");return(await this.sheetsApi.get(a,{baseURL:"",params:{id:this.spreadsheetId,format:e,...t&&{gid:t}},responseType:s?"stream":"arraybuffer"})).data}async downloadAsZippedHTML(e){return this._downloadAs("html",void 0,e)}async downloadAsHTML(e){return this._downloadAs("html",void 0,e)}async downloadAsXLSX(e=!1){return this._downloadAs("xlsx",void 0,e)}async downloadAsODS(e=!1){return this._downloadAs("ods",void 0,e)}async delete(){const e=await this.driveApi.delete("");return this._deleted=!0,e.data}async listPermissions(){return(await this.driveApi.request({method:"GET",url:"/permissions",params:{fields:"permissions(id,type,emailAddress,domain,role,displayName,photoLink,deleted)"}})).data.permissions}async setPublicAccessLevel(e){const t=await this.listPermissions(),s=find_js__default(t,a=>a.type==="anyone");if(e===!1){if(!s)return;await this.driveApi.request({method:"DELETE",url:`/permissions/${s.id}`})}else await this.driveApi.request({method:"POST",url:"/permissions",params:{},data:{role:e||"viewer",type:"anyone"}})}async share(e,t){let s,a;return e.includes("@")?s=e:a=e,(await this.driveApi.request({method:"POST",url:"/permissions",params:{...t?.emailMessage===!1&&{sendNotificationEmail:!1},...isString_js__default(t?.emailMessage)&&{emailMessage:t?.emailMessage},...t?.role==="owner"&&{transferOwnership:!0}},data:{role:t?.role||"writer",...s&&{type:t?.isGroup?"group":"user",emailAddress:s},...a&&{type:"domain",domain:a}}})).data}static async createNewSpreadsheetDocument(e,t){if(m(e)===AUTH_MODES.API_KEY)throw new Error("Cannot use api key only to create a new spreadsheet - it is only usable for read-only access of public docs");const s=await g(e),a=await p__default.request({method:"POST",url:u,paramsSerializer:axiosParamsSerializer,...s,data:{properties:t}}),n=new GoogleSpreadsheet(a.data.spreadsheetId,e);return n._spreadsheetUrl=a.data.spreadsheetUrl,n._rawProperties=a.data.properties,n._updateNamedRanges(a.data.namedRanges),each_js__default(a.data.sheets,o=>n._updateOrCreateSheet(o)),n}}exports.GoogleSpreadsheet=GoogleSpreadsheet,exports.GoogleSpreadsheetCell=GoogleSpreadsheetCell,exports.GoogleSpreadsheetCellErrorValue=GoogleSpreadsheetCellErrorValue,exports.GoogleSpreadsheetRow=GoogleSpreadsheetRow,exports.GoogleSpreadsheetWorksheet=GoogleSpreadsheetWorksheet,exports.axiosParamsSerializer=axiosParamsSerializer,exports.checkForDuplicateHeaders=checkForDuplicateHeaders,exports.columnToLetter=columnToLetter,exports.getFieldMask=getFieldMask,exports.letterToColumn=letterToColumn,exports.parseRangeA1=parseRangeA1,exports.toA1Range=toA1Range;