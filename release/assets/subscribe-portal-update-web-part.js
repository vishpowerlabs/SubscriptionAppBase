define("b6e1781d-7b91-48f8-aa1e-3bafe48cd49d_0.0.1", ["@microsoft/sp-core-library","@microsoft/sp-webpart-base","@microsoft/sp-http","@microsoft/sp-lodash-subset"], (__WEBPACK_EXTERNAL_MODULE__676__, __WEBPACK_EXTERNAL_MODULE__642__, __WEBPACK_EXTERNAL_MODULE__909__, __WEBPACK_EXTERNAL_MODULE__529__) => { return /******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 991:
/*!************************************************************************************!*\
  !*** ./lib/webparts/subscribePortalUpdate/SubscribePortalUpdateWebPart.module.css ***!
  \************************************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../../node_modules/@microsoft/sp-css-loader/node_modules/@microsoft/load-themed-styles/lib-es6/index.js */ 323);
// Imports


_node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__.loadStyles(".subscribePortalUpdate_a3571bf6{-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;font-family:Segoe UI,Tahoma,Geneva,Verdana,sans-serif;margin:0 auto;max-width:600px;padding:24px}.subscribePortalUpdate_a3571bf6 *{box-sizing:border-box}.subscribePortalUpdate_a3571bf6 .card_a3571bf6{background:\"[theme: white, default: #ffffff]\";border:1px solid;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,.04),0 8px 24px rgba(0,0,0,.08);overflow:hidden;transition:all .3s ease}.subscribePortalUpdate_a3571bf6 .card_a3571bf6:hover{box-shadow:0 4px 12px rgba(0,0,0,.06),0 12px 32px rgba(0,0,0,.12);transform:translateY(-2px)}.subscribePortalUpdate_a3571bf6 .cardHeader_a3571bf6{align-items:flex-start;background:\"[theme: themeLighterAlt, default: #eff6fc]\";border-bottom:1px solid;display:flex;padding:28px 32px}.subscribePortalUpdate_a3571bf6 .headerContent_a3571bf6{flex:1;margin:0;min-width:0;padding:0}.subscribePortalUpdate_a3571bf6 .title_a3571bf6{color:\"[theme: neutralPrimary, default: #333333]\";font-size:22px;font-weight:600;letter-spacing:-.02em;line-height:1.3;margin:0;padding:0}.subscribePortalUpdate_a3571bf6 .description_a3571bf6{color:\"[theme: neutralSecondary, default: #666666]\";font-size:14px;font-weight:400;line-height:1.6;margin:8px 0 0;padding:0}.subscribePortalUpdate_a3571bf6 .cardBody_a3571bf6{display:flex;flex-direction:column;gap:20px;padding:28px 32px}.subscribePortalUpdate_a3571bf6 .button_a3571bf6{align-items:center;border:none;border-radius:10px;cursor:pointer;display:flex;font-family:inherit;font-size:15px;font-weight:600;gap:10px;justify-content:center;letter-spacing:.01em;overflow:hidden;padding:14px 28px;position:relative;transition:all .2s ease;width:100%}.subscribePortalUpdate_a3571bf6 .button_a3571bf6:before{background:hsla(0,0%,100%,.2);border-radius:50%;content:\"\";height:0;left:50%;position:absolute;top:50%;transform:translate(-50%,-50%);transition:width .6s,height .6s;width:0}.subscribePortalUpdate_a3571bf6 .button_a3571bf6:hover:before{height:300px;width:300px}.subscribePortalUpdate_a3571bf6 .button_a3571bf6:active{transform:scale(.98)}.subscribePortalUpdate_a3571bf6 .button_a3571bf6:disabled{cursor:not-allowed;opacity:.6;transform:none}.subscribePortalUpdate_a3571bf6 .button_a3571bf6:disabled:hover:before{height:0;width:0}.subscribePortalUpdate_a3571bf6 .button_a3571bf6:focus-visible{outline:3px solid;outline-offset:2px}.subscribePortalUpdate_a3571bf6 .buttonPrimary_a3571bf6{background:\"[theme: themePrimary, default: #0078d4]\";box-shadow:0 2px 8px rgba(0,120,212,.3),0 4px 16px rgba(0,120,212,.2);color:#fff}.subscribePortalUpdate_a3571bf6 .buttonPrimary_a3571bf6:hover:not(:disabled){background:\"[theme: themeDarkAlt, default: #106ebe]\";box-shadow:0 4px 12px rgba(0,120,212,.4),0 6px 20px rgba(0,120,212,.3);transform:translateY(-2px)}.subscribePortalUpdate_a3571bf6 .buttonSecondary_a3571bf6{background:linear-gradient(135deg,#fff,#f8f9fa);border:2px solid;box-shadow:0 2px 6px rgba(0,0,0,.08);color:\"[theme: neutralPrimary, default: #333333]\"}.subscribePortalUpdate_a3571bf6 .buttonSecondary_a3571bf6:hover:not(:disabled){background:\"[theme: neutralLighter, default: #f4f4f4]\";border-color:\"[theme: themePrimary, default: #0078d4]\";box-shadow:0 4px 12px rgba(0,0,0,.12),0 0 0 1px \"[theme: themePrimary, default: #0078d4]\";color:\"[theme: themePrimary, default: #0078d4]\";transform:translateY(-2px)}.subscribePortalUpdate_a3571bf6 .buttonText_a3571bf6{position:relative;z-index:1}.subscribePortalUpdate_a3571bf6 .spinner_a3571bf6{animation:spin_a3571bf6 .8s linear infinite;border:3px solid hsla(0,0%,100%,.3);border-radius:50%;border-top-color:#fff;display:inline-block;height:18px;position:relative;width:18px;z-index:1}@keyframes spin_a3571bf6{to{transform:rotate(1turn)}}@media (max-width:640px){.subscribePortalUpdate_a3571bf6{padding:16px}.subscribePortalUpdate_a3571bf6 .cardHeader_a3571bf6{padding:20px}.subscribePortalUpdate_a3571bf6 .title_a3571bf6{font-size:19px}.subscribePortalUpdate_a3571bf6 .description_a3571bf6{font-size:13px}.subscribePortalUpdate_a3571bf6 .cardBody_a3571bf6{padding:20px}.subscribePortalUpdate_a3571bf6 .button_a3571bf6{font-size:14px;padding:12px 24px}}@media (max-width:480px){.subscribePortalUpdate_a3571bf6 .cardHeader_a3571bf6,.subscribePortalUpdate_a3571bf6 .headerContent_a3571bf6{text-align:center}}@media print{.subscribePortalUpdate_a3571bf6 .button_a3571bf6{display:none}.subscribePortalUpdate_a3571bf6 .card_a3571bf6{border:1px solid #000;box-shadow:none}}@media (prefers-reduced-motion:reduce){.subscribePortalUpdate_a3571bf6 *,.subscribePortalUpdate_a3571bf6 :after,.subscribePortalUpdate_a3571bf6 :before{animation-duration:0s!important;animation-iteration-count:1!important;transition-duration:0s!important}}@media (prefers-contrast:high){.subscribePortalUpdate_a3571bf6 .button_a3571bf6,.subscribePortalUpdate_a3571bf6 .card_a3571bf6{border:2px solid}}\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbImZpbGU6Ly8vd29ya3NwYWNlcy9TdWJzY3JpcHRpb25BcHBCYXNlL3NyYy93ZWJwYXJ0cy9zdWJzY3JpYmVQb3J0YWxVcGRhdGUvU3Vic2NyaWJlUG9ydGFsVXBkYXRlV2ViUGFydC5tb2R1bGUuc2NzcyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiQUFFQSxnQ0FLRSxrQ0FBQSxDQUNBLGlDQUFBLENBRkEscURBQUEsQ0FGQSxhQUFBLENBREEsZUFBQSxDQUVBLFlBR0EsQ0FFQSxrQ0FDRSxxQkFBQSxDQU1GLCtDQUNFLDZDQUFBLENBS0EsZ0JBQUEsQ0FKQSxrQkFBQSxDQUNBLCtEQUFBLENBSUEsZUFBQSxDQUNBLHVCQUFBLENBRUEscURBQ0UsaUVBQUEsQ0FHQSwwQkFBQSxDQUlKLHFEQUVFLHNCQUFBLENBRUEsdURBQUEsQ0FDQSx1QkFBQSxDQUpBLFlBQUEsQ0FFQSxpQkFFQSxDQUdGLHdEQUNFLE1BQUEsQ0FFQSxRQUFBLENBREEsV0FBQSxDQUVBLFNBQUEsQ0FHRixnREFLRSxpREFBQSxDQUZBLGNBQUEsQ0FDQSxlQUFBLENBR0EscUJBQUEsQ0FEQSxlQUFBLENBTEEsUUFBQSxDQUNBLFNBS0EsQ0FHRixzREFLRSxtREFBQSxDQUZBLGNBQUEsQ0FHQSxlQUFBLENBRkEsZUFBQSxDQUhBLGNBQUEsQ0FDQSxTQUlBLENBR0YsbURBRUUsWUFBQSxDQUNBLHFCQUFBLENBQ0EsUUFBQSxDQUhBLGlCQUdBLENBTUYsaURBVUUsa0JBQUEsQ0FMQSxXQUFBLENBQ0Esa0JBQUEsQ0FDQSxjQUFBLENBRUEsWUFBQSxDQU9BLG1CQUFBLENBYkEsY0FBQSxDQUNBLGVBQUEsQ0FRQSxRQUFBLENBREEsc0JBQUEsQ0FJQSxvQkFBQSxDQURBLGVBQUEsQ0FaQSxpQkFBQSxDQVdBLGlCQUFBLENBTEEsdUJBQUEsQ0FQQSxVQWVBLENBRUEsd0RBUUUsNkJBQUEsQ0FEQSxpQkFBQSxDQU5BLFVBQUEsQ0FLQSxRQUFBLENBRkEsUUFBQSxDQUZBLGlCQUFBLENBQ0EsT0FBQSxDQU1BLDhCQUFBLENBQ0EsK0JBQUEsQ0FMQSxPQUtBLENBR0YsOERBRUUsWUFBQSxDQURBLFdBQ0EsQ0FHRix3REFDRSxvQkFBQSxDQUdGLDBEQUNFLGtCQUFBLENBQ0EsVUFBQSxDQUNBLGNBQUEsQ0FFQSx1RUFFRSxRQUFBLENBREEsT0FDQSxDQUlKLCtEQUNFLGlCQUFBLENBQ0Esa0JBQUEsQ0FJSix3REFDRSxvREFBQSxDQUVBLHFFQUFBLENBREEsVUFDQSxDQUlBLDZFQUNFLG9EQUFBLENBQ0Esc0VBQUEsQ0FHQSwwQkFBQSxDQUlKLDBEQUNFLCtDQUFBLENBRUEsZ0JBQUEsQ0FDQSxvQ0FBQSxDQUZBLGlEQUVBLENBRUEsK0VBQ0Usc0RBQUEsQ0FDQSxzREFBQSxDQUVBLHlGQUFBLENBREEsK0NBQUEsQ0FJQSwwQkFBQSxDQUlKLHFEQUNFLGlCQUFBLENBQ0EsU0FBQSxDQU1GLGtEQU9FLDJDQUFBLENBRkEsbUNBQUEsQ0FDQSxpQkFBQSxDQURBLHFCQUFBLENBSkEsb0JBQUEsQ0FFQSxXQUFBLENBS0EsaUJBQUEsQ0FOQSxVQUFBLENBT0EsU0FBQSxDQU1GLHlCQUNFLEdBQ0UsdUJBQUEsQ0FBQSxDQU9KLHlCQXZNRixnQ0F3TUksWUFBQSxDQUVBLHFEQUNFLFlBQUEsQ0FHRixnREFDRSxjQUFBLENBR0Ysc0RBQ0UsY0FBQSxDQUdGLG1EQUNFLFlBQUEsQ0FHRixpREFFRSxjQUFBLENBREEsaUJBQ0EsQ0FBQSxDQUlKLHlCQUtFLDZHQUNFLGlCQUFBLENBQUEsQ0FPSixhQUNFLGlEQUNFLFlBQUEsQ0FHRiwrQ0FFRSxxQkFBQSxDQURBLGVBQ0EsQ0FBQSxDQU9KLHVDQUNFLGlIQUdFLCtCQUFBLENBQ0EscUNBQUEsQ0FDQSxnQ0FBQSxDQUFBLENBS0osK0JBS0UsZ0dBQ0UsZ0JBQUEsQ0FBQSIsImZpbGUiOiJTdWJzY3JpYmVQb3J0YWxVcGRhdGVXZWJQYXJ0Lm1vZHVsZS5jc3MifQ== */", true);

// Exports
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = ({
  subscribePortalUpdate_a3571bf6: "subscribePortalUpdate_a3571bf6",
  card_a3571bf6: "card_a3571bf6",
  cardHeader_a3571bf6: "cardHeader_a3571bf6",
  headerContent_a3571bf6: "headerContent_a3571bf6",
  title_a3571bf6: "title_a3571bf6",
  description_a3571bf6: "description_a3571bf6",
  cardBody_a3571bf6: "cardBody_a3571bf6",
  button_a3571bf6: "button_a3571bf6",
  buttonPrimary_a3571bf6: "buttonPrimary_a3571bf6",
  buttonSecondary_a3571bf6: "buttonSecondary_a3571bf6",
  buttonText_a3571bf6: "buttonText_a3571bf6",
  spinner_a3571bf6: "spinner_a3571bf6",
  spin_a3571bf6: "spin_a3571bf6"
});


/***/ }),

/***/ 451:
/*!****************************************************************************************!*\
  !*** ./lib/webparts/subscribePortalUpdate/SubscribePortalUpdateWebPart.module.scss.js ***!
  \****************************************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
__webpack_require__(/*! ./SubscribePortalUpdateWebPart.module.css */ 991);
var styles = {
    subscribePortalUpdate: 'subscribePortalUpdate_a3571bf6',
    card: 'card_a3571bf6',
    cardHeader: 'cardHeader_a3571bf6',
    headerContent: 'headerContent_a3571bf6',
    title: 'title_a3571bf6',
    description: 'description_a3571bf6',
    cardBody: 'cardBody_a3571bf6',
    button: 'button_a3571bf6',
    buttonPrimary: 'buttonPrimary_a3571bf6',
    buttonSecondary: 'buttonSecondary_a3571bf6',
    buttonText: 'buttonText_a3571bf6',
    spinner: 'spinner_a3571bf6',
    spin: 'spin_a3571bf6'
};
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (styles);


/***/ }),

/***/ 323:
/*!***********************************************************************************************************!*\
  !*** ./node_modules/@microsoft/sp-css-loader/node_modules/@microsoft/load-themed-styles/lib-es6/index.js ***!
  \***********************************************************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   ClearStyleOptions: () => (/* binding */ ClearStyleOptions),
/* harmony export */   Mode: () => (/* binding */ Mode),
/* harmony export */   clearStyles: () => (/* binding */ clearStyles),
/* harmony export */   configureLoadStyles: () => (/* binding */ configureLoadStyles),
/* harmony export */   configureRunMode: () => (/* binding */ configureRunMode),
/* harmony export */   detokenize: () => (/* binding */ detokenize),
/* harmony export */   flush: () => (/* binding */ flush),
/* harmony export */   loadStyles: () => (/* binding */ loadStyles),
/* harmony export */   loadTheme: () => (/* binding */ loadTheme),
/* harmony export */   splitStyles: () => (/* binding */ splitStyles)
/* harmony export */ });
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
/**
 * In sync mode, styles are registered as style elements synchronously with loadStyles() call.
 * In async mode, styles are buffered and registered as batch in async timer for performance purpose.
 */
var Mode;
(function (Mode) {
    Mode[Mode["sync"] = 0] = "sync";
    Mode[Mode["async"] = 1] = "async";
})(Mode || (Mode = {}));
/**
 * Themable styles and non-themable styles are tracked separately
 * Specify ClearStyleOptions when calling clearStyles API to specify which group of registered styles should be cleared.
 */
var ClearStyleOptions;
(function (ClearStyleOptions) {
    /** only themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["onlyThemable"] = 1] = "onlyThemable";
    /** only non-themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["onlyNonThemable"] = 2] = "onlyNonThemable";
    /** both themable and non-themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["all"] = 3] = "all";
})(ClearStyleOptions || (ClearStyleOptions = {}));
// Store the theming state in __themeState__ global scope for reuse in the case of duplicate
// load-themed-styles hosted on the page.
var _root = typeof window === 'undefined' ? __webpack_require__.g : window; // eslint-disable-line @typescript-eslint/no-explicit-any
// Nonce string to inject into script tag if one provided. This is used in CSP (Content Security Policy).
var _styleNonce = _root && _root.CSPSettings && _root.CSPSettings.nonce;
var _themeState = initializeThemeState();
/**
 * Matches theming tokens. For example, "[theme: themeSlotName, default: #FFF]" (including the quotes).
 */
var _themeTokenRegex = /[\'\"]\[theme:\s*(\w+)\s*(?:\,\s*default:\s*([\\"\']?[\.\,\(\)\#\-\s\w]*[\.\,\(\)\#\-\w][\"\']?))?\s*\][\'\"]/g;
var now = function () {
    return typeof performance !== 'undefined' && !!performance.now ? performance.now() : Date.now();
};
function measure(func) {
    var start = now();
    func();
    var end = now();
    _themeState.perf.duration += end - start;
}
/**
 * initialize global state object
 */
function initializeThemeState() {
    var state = _root.__themeState__ || {
        theme: undefined,
        lastStyleElement: undefined,
        registeredStyles: []
    };
    if (!state.runState) {
        state = __assign(__assign({}, state), { perf: {
                count: 0,
                duration: 0
            }, runState: {
                flushTimer: 0,
                mode: Mode.sync,
                buffer: []
            } });
    }
    if (!state.registeredThemableStyles) {
        state = __assign(__assign({}, state), { registeredThemableStyles: [] });
    }
    _root.__themeState__ = state;
    return state;
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load
 * event is fired.
 * @param {string | ThemableArray} styles Themable style text to register.
 * @param {boolean} loadAsync When true, always load styles in async mode, irrespective of current sync mode.
 */
function loadStyles(styles, loadAsync) {
    if (loadAsync === void 0) { loadAsync = false; }
    measure(function () {
        var styleParts = Array.isArray(styles) ? styles : splitStyles(styles);
        var _a = _themeState.runState, mode = _a.mode, buffer = _a.buffer, flushTimer = _a.flushTimer;
        if (loadAsync || mode === Mode.async) {
            buffer.push(styleParts);
            if (!flushTimer) {
                _themeState.runState.flushTimer = asyncLoadStyles();
            }
        }
        else {
            applyThemableStyles(styleParts);
        }
    });
}
/**
 * Allows for customizable loadStyles logic. e.g. for server side rendering application
 * @param {(processedStyles: string, rawStyles?: string | ThemableArray) => void}
 * a loadStyles callback that gets called when styles are loaded or reloaded
 */
function configureLoadStyles(loadStylesFn) {
    _themeState.loadStyles = loadStylesFn;
}
/**
 * Configure run mode of load-themable-styles
 * @param mode load-themable-styles run mode, async or sync
 */
function configureRunMode(mode) {
    _themeState.runState.mode = mode;
}
/**
 * external code can call flush to synchronously force processing of currently buffered styles
 */
function flush() {
    measure(function () {
        var styleArrays = _themeState.runState.buffer.slice();
        _themeState.runState.buffer = [];
        var mergedStyleArray = [].concat.apply([], styleArrays);
        if (mergedStyleArray.length > 0) {
            applyThemableStyles(mergedStyleArray);
        }
    });
}
/**
 * register async loadStyles
 */
function asyncLoadStyles() {
    // Use "self" to distinguish conflicting global typings for setTimeout() from lib.dom.d.ts vs Jest's @types/node
    // https://github.com/jestjs/jest/issues/14418
    return self.setTimeout(function () {
        _themeState.runState.flushTimer = 0;
        flush();
    }, 0);
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load event
 * is fired.
 * @param {string} styleText Style to register.
 * @param {IStyleRecord} styleRecord Existing style record to re-apply.
 */
function applyThemableStyles(stylesArray, styleRecord) {
    if (_themeState.loadStyles) {
        _themeState.loadStyles(resolveThemableArray(stylesArray).styleString, stylesArray);
    }
    else {
        registerStyles(stylesArray);
    }
}
/**
 * Registers a set theme tokens to find and replace. If styles were already registered, they will be
 * replaced.
 * @param {theme} theme JSON object of theme tokens to values.
 */
function loadTheme(theme) {
    _themeState.theme = theme;
    // reload styles.
    reloadStyles();
}
/**
 * Clear already registered style elements and style records in theme_State object
 * @param option - specify which group of registered styles should be cleared.
 * Default to be both themable and non-themable styles will be cleared
 */
function clearStyles(option) {
    if (option === void 0) { option = ClearStyleOptions.all; }
    if (option === ClearStyleOptions.all || option === ClearStyleOptions.onlyNonThemable) {
        clearStylesInternal(_themeState.registeredStyles);
        _themeState.registeredStyles = [];
    }
    if (option === ClearStyleOptions.all || option === ClearStyleOptions.onlyThemable) {
        clearStylesInternal(_themeState.registeredThemableStyles);
        _themeState.registeredThemableStyles = [];
    }
}
function clearStylesInternal(records) {
    records.forEach(function (styleRecord) {
        var styleElement = styleRecord && styleRecord.styleElement;
        if (styleElement && styleElement.parentElement) {
            styleElement.parentElement.removeChild(styleElement);
        }
    });
}
/**
 * Reloads styles.
 */
function reloadStyles() {
    if (_themeState.theme) {
        var themableStyles = [];
        for (var _i = 0, _a = _themeState.registeredThemableStyles; _i < _a.length; _i++) {
            var styleRecord = _a[_i];
            themableStyles.push(styleRecord.themableStyle);
        }
        if (themableStyles.length > 0) {
            clearStyles(ClearStyleOptions.onlyThemable);
            applyThemableStyles([].concat.apply([], themableStyles));
        }
    }
}
/**
 * Find theme tokens and replaces them with provided theme values.
 * @param {string} styles Tokenized styles to fix.
 */
function detokenize(styles) {
    if (styles) {
        styles = resolveThemableArray(splitStyles(styles)).styleString;
    }
    return styles;
}
/**
 * Resolves ThemingInstruction objects in an array and joins the result into a string.
 * @param {ThemableArray} splitStyleArray ThemableArray to resolve and join.
 */
function resolveThemableArray(splitStyleArray) {
    var theme = _themeState.theme;
    var themable = false;
    // Resolve the array of theming instructions to an array of strings.
    // Then join the array to produce the final CSS string.
    var resolvedArray = (splitStyleArray || []).map(function (currentValue) {
        var themeSlot = currentValue.theme;
        if (themeSlot) {
            themable = true;
            // A theming annotation. Resolve it.
            var themedValue = theme ? theme[themeSlot] : undefined;
            var defaultValue = currentValue.defaultValue || 'inherit';
            // Warn to console if we hit an unthemed value even when themes are provided, but only if "DEBUG" is true.
            // Allow the themedValue to be undefined to explicitly request the default value.
            if (theme &&
                !themedValue &&
                console &&
                !(themeSlot in theme) &&
                "boolean" !== 'undefined' &&
                true) {
                // eslint-disable-next-line no-console
                console.warn("Theming value not provided for \"".concat(themeSlot, "\". Falling back to \"").concat(defaultValue, "\"."));
            }
            return themedValue || defaultValue;
        }
        else {
            // A non-themable string. Preserve it.
            return currentValue.rawString;
        }
    });
    return {
        styleString: resolvedArray.join(''),
        themable: themable
    };
}
/**
 * Split tokenized CSS into an array of strings and theme specification objects
 * @param {string} styles Tokenized styles to split.
 */
function splitStyles(styles) {
    var result = [];
    if (styles) {
        var pos = 0; // Current position in styles.
        var tokenMatch = void 0;
        while ((tokenMatch = _themeTokenRegex.exec(styles))) {
            var matchIndex = tokenMatch.index;
            if (matchIndex > pos) {
                result.push({
                    rawString: styles.substring(pos, matchIndex)
                });
            }
            result.push({
                theme: tokenMatch[1],
                defaultValue: tokenMatch[2] // May be undefined
            });
            // index of the first character after the current match
            pos = _themeTokenRegex.lastIndex;
        }
        // Push the rest of the string after the last match.
        result.push({
            rawString: styles.substring(pos)
        });
    }
    return result;
}
/**
 * Registers a set of style text. If it is registered too early, we will register it when the
 * window.load event is fired.
 * @param {ThemableArray} styleArray Array of IThemingInstruction objects to register.
 * @param {IStyleRecord} styleRecord May specify a style Element to update.
 */
function registerStyles(styleArray) {
    if (typeof document === 'undefined') {
        return;
    }
    var head = document.getElementsByTagName('head')[0];
    var styleElement = document.createElement('style');
    var _a = resolveThemableArray(styleArray), styleString = _a.styleString, themable = _a.themable;
    styleElement.setAttribute('data-load-themed-styles', 'true');
    if (_styleNonce) {
        styleElement.setAttribute('nonce', _styleNonce);
    }
    styleElement.appendChild(document.createTextNode(styleString));
    _themeState.perf.count++;
    head.appendChild(styleElement);
    var ev = document.createEvent('HTMLEvents');
    ev.initEvent('styleinsert', true /* bubbleEvent */, false /* cancelable */);
    ev.args = {
        newStyle: styleElement
    };
    document.dispatchEvent(ev);
    var record = {
        styleElement: styleElement,
        themableStyle: styleArray
    };
    if (themable) {
        _themeState.registeredThemableStyles.push(record);
    }
    else {
        _themeState.registeredStyles.push(record);
    }
}


/***/ }),

/***/ 676:
/*!*********************************************!*\
  !*** external "@microsoft/sp-core-library" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__676__;

/***/ }),

/***/ 909:
/*!*************************************!*\
  !*** external "@microsoft/sp-http" ***!
  \*************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__909__;

/***/ }),

/***/ 529:
/*!**********************************************!*\
  !*** external "@microsoft/sp-lodash-subset" ***!
  \**********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__529__;

/***/ }),

/***/ 642:
/*!*********************************************!*\
  !*** external "@microsoft/sp-webpart-base" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__642__;

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	(() => {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
(() => {
/*!****************************************************************************!*\
  !*** ./lib/webparts/subscribePortalUpdate/SubscribePortalUpdateWebPart.js ***!
  \****************************************************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-core-library */ 676);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @microsoft/sp-webpart-base */ 642);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/sp-http */ 909);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./SubscribePortalUpdateWebPart.module.scss */ 451);
/* harmony import */ var _microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @microsoft/sp-lodash-subset */ 529);
/* harmony import */ var _microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_4___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_4__);
var __extends = (undefined && undefined.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};





var SubscribePortalUpdateWebPart = /** @class */ (function (_super) {
    __extends(SubscribePortalUpdateWebPart, _super);
    function SubscribePortalUpdateWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isSubscribed = false;
        _this._isLoading = false;
        _this._currentUserEmail = '';
        _this._alreadyInitialized = false;
        _this._productNameOptions = [];
        _this._productNameOptionsLoaded = false;
        return _this;
    }
    // ----------------- Helpers / defaults -----------------
    SubscribePortalUpdateWebPart.prototype._ensureDefaultProperties = function () {
        if (!this.properties.ListName || this.properties.ListName.trim().length === 0) {
            this.properties.ListName = 'subscriptionlist';
        }
        if (!this.properties.PrefixText || this.properties.PrefixText.trim().length === 0) {
            this.properties.PrefixText = 'Subscribe to';
        }
        if (!this.properties.SuffixText || this.properties.SuffixText.trim().length === 0) {
            this.properties.SuffixText = 'updates';
        }
        if (!this.properties.ProductNameColumnName || this.properties.ProductNameColumnName.trim().length === 0) {
            this.properties.ProductNameColumnName = 'SubscribedFor';
        }
        if (!this.properties.HeaderFontSize || this.properties.HeaderFontSize <= 0) {
            this.properties.HeaderFontSize = 22;
        }
        if (!this.properties.DescriptionFontSize || this.properties.DescriptionFontSize <= 0) {
            this.properties.DescriptionFontSize = 14;
        }
        if (!this.properties.ButtonFontSize || this.properties.ButtonFontSize <= 0) {
            this.properties.ButtonFontSize = 15;
        }
    };
    SubscribePortalUpdateWebPart.prototype._logDebug = function (message, extra) {
        // eslint-disable-next-line no-console
        console.log("[SubscribePortalUpdate] ".concat(message), extra !== null && extra !== void 0 ? extra : '');
    };
    // ----------------- Lifecycle -----------------
    SubscribePortalUpdateWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                this._ensureDefaultProperties();
                this._logDebug('onInit', {
                    listName: this.properties.ListName,
                    productName: this.properties.ProductName
                });
                return [2 /*return*/, _super.prototype.onInit.call(this)];
            });
        });
    };
    SubscribePortalUpdateWebPart.prototype.render = function () {
        this._ensureDefaultProperties();
        this.domElement.innerHTML = this._getHtml();
        this._wireEvents();
        if (!this._alreadyInitialized) {
            this._alreadyInitialized = true;
            void this._loadInitialStatus();
        }
    };
    SubscribePortalUpdateWebPart.prototype._loadInitialStatus = function () {
        return __awaiter(this, void 0, void 0, function () {
            var err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, 4, 5]);
                        this._setLoading(true);
                        return [4 /*yield*/, this._ensureUserEmail()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this._checkSubscription()];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 5];
                    case 3:
                        err_1 = _a.sent();
                        this._logDebug('Error checking subscription status', err_1);
                        return [3 /*break*/, 5];
                    case 4:
                        this._setLoading(false);
                        return [7 /*endfinally*/];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    // ----------------- User / context -----------------
    SubscribePortalUpdateWebPart.prototype._ensureUserEmail = function () {
        var _a;
        return __awaiter(this, void 0, void 0, function () {
            var url, resp, text, json;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (this._currentUserEmail) {
                            return [2 /*return*/];
                        }
                        if ((_a = this.context.pageContext.user) === null || _a === void 0 ? void 0 : _a.email) {
                            this._currentUserEmail = this.context.pageContext.user.email;
                            this._logDebug('Using pageContext user email', this._currentUserEmail);
                            return [2 /*return*/];
                        }
                        url = "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/currentuser");
                        this._logDebug('GET current user', url);
                        return [4 /*yield*/, this.context.spHttpClient.get(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata.metadata=none'
                                }
                            })];
                    case 1:
                        resp = _b.sent();
                        if (!!resp.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, resp.text()];
                    case 2:
                        text = _b.sent();
                        this._logDebug('Error response from currentuser', { status: resp.status, body: text });
                        throw new Error("Unable to get current user. Status: ".concat(resp.status, ". Body: ").concat(text));
                    case 3: return [4 /*yield*/, resp.json()];
                    case 4:
                        json = _b.sent();
                        this._currentUserEmail = json.Email;
                        this._logDebug('Resolved current user email', this._currentUserEmail);
                        return [2 /*return*/];
                }
            });
        });
    };
    // ----------------- Subscription status -----------------
    SubscribePortalUpdateWebPart.prototype._checkSubscription = function () {
        return __awaiter(this, void 0, void 0, function () {
            var item, arr;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._getSubscriptionItem()];
                    case 1:
                        item = _a.sent();
                        this._logDebug('Loaded subscription item', item);
                        if (!item || !this.properties.ProductName) {
                            this._isSubscribed = false;
                        }
                        else {
                            arr = item.SubscribedFor || [];
                            this._isSubscribed = arr.indexOf(this.properties.ProductName) > -1;
                        }
                        this.render();
                        return [2 /*return*/];
                }
            });
        });
    };
    // ----------------- Subscribe / Unsubscribe -----------------
    SubscribePortalUpdateWebPart.prototype._subscribe = function () {
        return __awaiter(this, void 0, void 0, function () {
            var product, item, current, updated, err_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this._setLoading(true);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 8, , 9]);
                        product = this.properties.ProductName;
                        if (!product) {
                            throw new Error('ProductName is not configured.');
                        }
                        return [4 /*yield*/, this._getSubscriptionItem()];
                    case 2:
                        item = _a.sent();
                        if (!!item) return [3 /*break*/, 4];
                        this._logDebug('No existing subscription item, creating new');
                        return [4 /*yield*/, this._createSubscriptionItem(product)];
                    case 3:
                        _a.sent();
                        return [3 /*break*/, 7];
                    case 4:
                        current = item.SubscribedFor || [];
                        if (!(current.indexOf(product) === -1)) return [3 /*break*/, 6];
                        updated = current.slice();
                        updated.push(product);
                        this._logDebug('Updating existing item with new product', { id: item.Id, updated: updated });
                        return [4 /*yield*/, this._updateSubscriptionItem(item.Id, updated)];
                    case 5:
                        _a.sent();
                        return [3 /*break*/, 7];
                    case 6:
                        this._isSubscribed = true;
                        this.render();
                        this._setLoading(false);
                        return [2 /*return*/];
                    case 7:
                        this._isSubscribed = true;
                        this.render();
                        return [3 /*break*/, 9];
                    case 8:
                        err_2 = _a.sent();
                        this._logDebug('Error subscribing', err_2);
                        return [3 /*break*/, 9];
                    case 9:
                        this._setLoading(false);
                        return [2 /*return*/];
                }
            });
        });
    };
    SubscribePortalUpdateWebPart.prototype._unsubscribe = function () {
        return __awaiter(this, void 0, void 0, function () {
            var product_1, item, current, remaining, err_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this._setLoading(true);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        product_1 = this.properties.ProductName;
                        if (!product_1) {
                            throw new Error('ProductName is not configured.');
                        }
                        return [4 /*yield*/, this._getSubscriptionItem()];
                    case 2:
                        item = _a.sent();
                        if (!item) {
                            this._isSubscribed = false;
                            this.render();
                            this._setLoading(false);
                            return [2 /*return*/];
                        }
                        current = item.SubscribedFor || [];
                        remaining = current.filter(function (x) { return x !== product_1; });
                        this._logDebug('Updating item for unsubscribe', {
                            id: item.Id,
                            current: current,
                            remaining: remaining
                        });
                        return [4 /*yield*/, this._updateSubscriptionItem(item.Id, remaining)];
                    case 3:
                        _a.sent();
                        this._isSubscribed = false;
                        this.render();
                        return [3 /*break*/, 5];
                    case 4:
                        err_3 = _a.sent();
                        this._logDebug('Error unsubscribing', err_3);
                        return [3 /*break*/, 5];
                    case 5:
                        this._setLoading(false);
                        return [2 /*return*/];
                }
            });
        });
    };
    // ----------------- REST helpers -----------------
    SubscribePortalUpdateWebPart.prototype._getSubscriptionItem = function () {
        return __awaiter(this, void 0, void 0, function () {
            var list, email, url, resp, text, json, row, raw, subscribedFor, result;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        list = this.properties.ListName;
                        if (!list) {
                            throw new Error('ListName is not configured.');
                        }
                        email = this._currentUserEmail;
                        url = "".concat(this.context.pageContext.web.absoluteUrl) +
                            "/_api/web/lists/getbytitle('".concat(list, "')/items") +
                            "?$filter=SubscribedBy eq '".concat(email, "'") +
                            "&$select=Id,Title,SubscribedBy,SubscribedFor,Subscribe,SubscribeOn";
                        this._logDebug('GET subscription item', url);
                        return [4 /*yield*/, this.context.spHttpClient.get(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata.metadata=none'
                                }
                            })];
                    case 1:
                        resp = _a.sent();
                        if (!!resp.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, resp.text()];
                    case 2:
                        text = _a.sent();
                        this._logDebug('Error from GET subscription item', { status: resp.status, body: text });
                        if (resp.status === 404) {
                            throw new Error('List not found. Body: ' + text);
                        }
                        throw new Error("Failed to get subscription item. Status: ".concat(resp.status, ". Body: ").concat(text));
                    case 3: return [4 /*yield*/, resp.json()];
                    case 4:
                        json = _a.sent();
                        row = (json.value && json.value[0]) || null;
                        if (!row) {
                            return [2 /*return*/, null];
                        }
                        raw = row.SubscribedFor;
                        subscribedFor = [];
                        if (Array.isArray(raw)) {
                            subscribedFor = raw;
                        }
                        else if (raw && Array.isArray(raw.results)) {
                            subscribedFor = raw.results;
                        }
                        result = {
                            Id: row.Id,
                            Title: row.Title,
                            SubscribedBy: row.SubscribedBy,
                            Subscribe: row.Subscribe,
                            SubscribeOn: row.SubscribeOn,
                            SubscribedFor: subscribedFor
                        };
                        this._logDebug('Parsed subscription item', result);
                        return [2 /*return*/, result];
                }
            });
        });
    };
    SubscribePortalUpdateWebPart.prototype._createSubscriptionItem = function (product) {
        return __awaiter(this, void 0, void 0, function () {
            var list, body, url, resp, text, json;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        list = this.properties.ListName;
                        if (!list) {
                            throw new Error('ListName is not configured.');
                        }
                        body = {
                            Title: this._currentUserEmail,
                            SubscribedBy: this._currentUserEmail,
                            Subscribe: 'Yes',
                            SubscribeOn: new Date().toISOString(),
                            SubscribedFor: [product]
                        };
                        url = "".concat(this.context.pageContext.web.absoluteUrl) +
                            "/_api/web/lists/getbytitle('".concat(list, "')/items");
                        this._logDebug('POST create subscription item', { url: url, body: body });
                        return [4 /*yield*/, this.context.spHttpClient.post(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata.metadata=none',
                                    'Content-Type': 'application/json;odata.metadata=none'
                                },
                                body: JSON.stringify(body)
                            })];
                    case 1:
                        resp = _a.sent();
                        if (!!resp.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, resp.text()];
                    case 2:
                        text = _a.sent();
                        this._logDebug('Error from POST create item', { status: resp.status, body: text });
                        throw new Error("Failed to create subscription item. Status: ".concat(resp.status, ". Body: ").concat(text));
                    case 3: return [4 /*yield*/, resp.json()];
                    case 4:
                        json = _a.sent();
                        this._logDebug('Created subscription item', json);
                        return [2 /*return*/];
                }
            });
        });
    };
    SubscribePortalUpdateWebPart.prototype._updateSubscriptionItem = function (id, products) {
        return __awaiter(this, void 0, void 0, function () {
            var list, body, url, resp, text;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        list = this.properties.ListName;
                        if (!list) {
                            throw new Error('ListName is not configured.');
                        }
                        body = {
                            SubscribeOn: new Date().toISOString(),
                            Subscribe: products.length > 0 ? 'Yes' : 'No',
                            SubscribedFor: products
                        };
                        url = "".concat(this.context.pageContext.web.absoluteUrl) +
                            "/_api/web/lists/getbytitle('".concat(list, "')/items(").concat(id, ")");
                        this._logDebug('POST update subscription item', { url: url, body: body });
                        return [4 /*yield*/, this.context.spHttpClient.post(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata.metadata=none',
                                    'Content-Type': 'application/json;odata.metadata=none',
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'MERGE'
                                },
                                body: JSON.stringify(body)
                            })];
                    case 1:
                        resp = _a.sent();
                        if (!!resp.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, resp.text()];
                    case 2:
                        text = _a.sent();
                        this._logDebug('Error from POST update item', { status: resp.status, body: text });
                        throw new Error("Failed to update subscription item. Status: ".concat(resp.status, ". Body: ").concat(text));
                    case 3:
                        this._logDebug('Update subscription item OK');
                        return [2 /*return*/];
                }
            });
        });
    };
    // ----------------- Load Product Name Choices -----------------
    SubscribePortalUpdateWebPart.prototype._loadProductNameChoices = function () {
        return __awaiter(this, void 0, void 0, function () {
            var list, columnName, url, resp, text, json, field, choices, options, err_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        list = this.properties.ListName;
                        columnName = this.properties.ProductNameColumnName;
                        if (!list || !columnName) {
                            this._logDebug('List name or column name not configured', { list: list, columnName: columnName });
                            return [2 /*return*/, []];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 6, , 7]);
                        url = "".concat(this.context.pageContext.web.absoluteUrl) +
                            "/_api/web/lists/getbytitle('".concat(list, "')/fields") +
                            "?$filter=EntityPropertyName eq '".concat(columnName, "'") +
                            "&$select=Choices,TypeAsString";
                        this._logDebug('GET field choices', url);
                        return [4 /*yield*/, this.context.spHttpClient.get(url, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_2__.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata.metadata=minimal'
                                }
                            })];
                    case 2:
                        resp = _a.sent();
                        if (!!resp.ok) return [3 /*break*/, 4];
                        return [4 /*yield*/, resp.text()];
                    case 3:
                        text = _a.sent();
                        this._logDebug('Error from GET field choices', { status: resp.status, body: text });
                        throw new Error("Failed to get field choices. Status: ".concat(resp.status));
                    case 4: return [4 /*yield*/, resp.json()];
                    case 5:
                        json = _a.sent();
                        if (!json.value || json.value.length === 0) {
                            this._logDebug('Field not found', columnName);
                            return [2 /*return*/, []];
                        }
                        field = json.value[0];
                        choices = field.Choices || [];
                        this._logDebug('Loaded choices', choices);
                        options = choices.map(function (choice) { return ({
                            key: choice,
                            text: choice
                        }); });
                        return [2 /*return*/, options];
                    case 6:
                        err_4 = _a.sent();
                        this._logDebug('Error loading product name choices', err_4);
                        return [2 /*return*/, []];
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    // ----------------- UI helpers -----------------
    SubscribePortalUpdateWebPart.prototype._setLoading = function (flag) {
        this._isLoading = flag;
        this.render();
    };
    SubscribePortalUpdateWebPart.prototype._getHtml = function () {
        var product = this.properties.ProductName || '';
        var prefix = this.properties.PrefixText || 'Subscribe to';
        var suffix = this.properties.SuffixText || 'updates';
        var buttonText = this._isSubscribed
            ? "Unsubscribe from ".concat(product)
            : "".concat(prefix, " ").concat(product, " ").concat(suffix).replace(/\s+/g, ' ').trim();
        var btnClass = this._isSubscribed ? _SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].buttonSecondary : _SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].buttonPrimary;
        return "\n      <div class=\"".concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].subscribePortalUpdate, "\">\n        <div class=\"").concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].card, "\">\n          <div class=\"").concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].cardHeader, "\">\n            <div class=\"").concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].headerContent, "\">\n              <h3 class=\"").concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].title, "\" style=\"font-size: ").concat(this.properties.HeaderFontSize, "px;\">\n                ").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_4__.escape)(product || 'Product Updates'), "\n              </h3>\n              ").concat(this.properties.Description ? "\n                <p class=\"".concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].description, "\" style=\"font-size: ").concat(this.properties.DescriptionFontSize, "px;\">\n                  ").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_4__.escape)(this.properties.Description), "\n                </p>\n              ") : '', "\n            </div>\n          </div>\n\n          <div class=\"").concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].cardBody, "\">\n            <button \n              id=\"toggleBtn\"\n              type=\"button\"\n              class=\"").concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].button, " ").concat(btnClass, "\"\n              ").concat(this._isLoading ? 'disabled' : '', "\n              aria-busy=\"").concat(this._isLoading, "\"\n              style=\"font-size: ").concat(this.properties.ButtonFontSize, "px;\"\n            >\n              ").concat(this._isLoading ? "\n                <span class=\"".concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].spinner, "\"></span>\n                <span>Processing...</span>\n              ") : "\n                <span class=\"".concat(_SubscribePortalUpdateWebPart_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].buttonText, "\">").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_4__.escape)(buttonText), "</span>\n              "), "\n            </button>\n          </div>\n        </div>\n      </div>\n    ");
    };
    SubscribePortalUpdateWebPart.prototype._wireEvents = function () {
        var _this = this;
        var btn = this.domElement.querySelector('#toggleBtn');
        if (btn) {
            btn.addEventListener('click', function () {
                if (_this._isLoading) {
                    return;
                }
                if (_this._isSubscribed) {
                    void _this._unsubscribe();
                }
                else {
                    void _this._subscribe();
                }
            });
        }
    };
    Object.defineProperty(SubscribePortalUpdateWebPart.prototype, "dataVersion", {
        // ----------------- SPFx plumbing -----------------
        get: function () {
            return _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(SubscribePortalUpdateWebPart.prototype, "disableReactivePropertyChanges", {
        get: function () {
            return false;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(SubscribePortalUpdateWebPart.prototype, "supportsThemeVariants", {
        // Support section backgrounds
        get: function () {
            return true;
        },
        enumerable: false,
        configurable: true
    });
    SubscribePortalUpdateWebPart.prototype.onPropertyPaneConfigurationStart = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!!this._productNameOptionsLoaded) return [3 /*break*/, 2];
                        this.context.propertyPane.refresh();
                        _a = this;
                        return [4 /*yield*/, this._loadProductNameChoices()];
                    case 1:
                        _a._productNameOptions = _b.sent();
                        this._productNameOptionsLoaded = true;
                        this.context.propertyPane.refresh();
                        _b.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    SubscribePortalUpdateWebPart.prototype.onPropertyPaneFieldChanged = function (propertyPath, oldValue, newValue) {
        var _this = this;
        // If ListName or ProductNameColumnName changes, reload the choices
        if (propertyPath === 'ListName' || propertyPath === 'ProductNameColumnName') {
            this._productNameOptionsLoaded = false;
            this._productNameOptions = [];
            // Reload choices
            void this._loadProductNameChoices().then(function (options) {
                _this._productNameOptions = options;
                _this._productNameOptionsLoaded = true;
                _this.context.propertyPane.refresh();
            });
        }
        _super.prototype.onPropertyPaneFieldChanged.call(this, propertyPath, oldValue, newValue);
    };
    SubscribePortalUpdateWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: { description: 'Subscribe Portal Update Settings' },
                    groups: [
                        {
                            groupName: 'List Configuration',
                            groupFields: [
                                (0,_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('ListName', {
                                    label: 'List Name',
                                    description: 'Default: "subscriptionlist"'
                                }),
                                (0,_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('ProductNameColumnName', {
                                    label: 'Product Name Column',
                                    description: 'The internal name of the multi-choice column (Default: "SubscribedFor")'
                                })
                            ]
                        },
                        {
                            groupName: 'Product Selection',
                            groupFields: [
                                (0,_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneDropdown)('ProductName', {
                                    label: 'Product Name',
                                    options: this._productNameOptions,
                                    disabled: !this._productNameOptionsLoaded || this._productNameOptions.length === 0
                                })
                            ]
                        },
                        {
                            groupName: 'Display Settings',
                            groupFields: [
                                (0,_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('Description', {
                                    label: 'Description',
                                    multiline: true
                                }),
                                (0,_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('PrefixText', {
                                    label: 'Button Prefix Text',
                                    description: 'Default: "Subscribe to"'
                                }),
                                (0,_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('SuffixText', {
                                    label: 'Button Suffix Text',
                                    description: 'Default: "updates"'
                                })
                            ]
                        },
                        {
                            groupName: 'Font Sizes',
                            groupFields: [
                                (0,_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneSlider)('HeaderFontSize', {
                                    label: 'Header Font Size (px)',
                                    min: 12,
                                    max: 48,
                                    step: 1,
                                    value: this.properties.HeaderFontSize || 22,
                                    showValue: true
                                }),
                                (0,_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneSlider)('DescriptionFontSize', {
                                    label: 'Description Font Size (px)',
                                    min: 10,
                                    max: 24,
                                    step: 1,
                                    value: this.properties.DescriptionFontSize || 14,
                                    showValue: true
                                }),
                                (0,_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneSlider)('ButtonFontSize', {
                                    label: 'Button Font Size (px)',
                                    min: 12,
                                    max: 24,
                                    step: 1,
                                    value: this.properties.ButtonFontSize || 15,
                                    showValue: true
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return SubscribePortalUpdateWebPart;
}(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_1__.BaseClientSideWebPart));
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (SubscribePortalUpdateWebPart);
//vishnu

})();

/******/ 	return __webpack_exports__;
/******/ })()
;
});;
//# sourceMappingURL=subscribe-portal-update-web-part.js.map