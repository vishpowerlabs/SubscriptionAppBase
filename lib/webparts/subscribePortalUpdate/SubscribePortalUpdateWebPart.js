var __extends = (this && this.__extends) || (function () {
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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
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
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneTextField, PropertyPaneDropdown, PropertyPaneSlider } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './SubscribePortalUpdateWebPart.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
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
                        return [4 /*yield*/, this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
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
                        return [4 /*yield*/, this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
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
                        return [4 /*yield*/, this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
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
                        return [4 /*yield*/, this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
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
                        return [4 /*yield*/, this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
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
        var btnClass = this._isSubscribed ? styles.buttonSecondary : styles.buttonPrimary;
        return "\n      <div class=\"".concat(styles.subscribePortalUpdate, "\">\n        <div class=\"").concat(styles.card, "\">\n          <div class=\"").concat(styles.cardHeader, "\">\n            <div class=\"").concat(styles.headerContent, "\">\n              <h3 class=\"").concat(styles.title, "\" style=\"font-size: ").concat(this.properties.HeaderFontSize, "px;\">\n                ").concat(escape(product || 'Product Updates'), "\n              </h3>\n              ").concat(this.properties.Description ? "\n                <p class=\"".concat(styles.description, "\" style=\"font-size: ").concat(this.properties.DescriptionFontSize, "px;\">\n                  ").concat(escape(this.properties.Description), "\n                </p>\n              ") : '', "\n            </div>\n          </div>\n\n          <div class=\"").concat(styles.cardBody, "\">\n            <button \n              id=\"toggleBtn\"\n              type=\"button\"\n              class=\"").concat(styles.button, " ").concat(btnClass, "\"\n              ").concat(this._isLoading ? 'disabled' : '', "\n              aria-busy=\"").concat(this._isLoading, "\"\n              style=\"font-size: ").concat(this.properties.ButtonFontSize, "px;\"\n            >\n              ").concat(this._isLoading ? "\n                <span class=\"".concat(styles.spinner, "\"></span>\n                <span>Processing...</span>\n              ") : "\n                <span class=\"".concat(styles.buttonText, "\">").concat(escape(buttonText), "</span>\n              "), "\n            </button>\n          </div>\n        </div>\n      </div>\n    ");
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
            return Version.parse('1.0');
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
                                PropertyPaneTextField('ListName', {
                                    label: 'List Name',
                                    description: 'Default: "subscriptionlist"'
                                }),
                                PropertyPaneTextField('ProductNameColumnName', {
                                    label: 'Product Name Column',
                                    description: 'The internal name of the multi-choice column (Default: "SubscribedFor")'
                                })
                            ]
                        },
                        {
                            groupName: 'Product Selection',
                            groupFields: [
                                PropertyPaneDropdown('ProductName', {
                                    label: 'Product Name',
                                    options: this._productNameOptions,
                                    disabled: !this._productNameOptionsLoaded || this._productNameOptions.length === 0
                                })
                            ]
                        },
                        {
                            groupName: 'Display Settings',
                            groupFields: [
                                PropertyPaneTextField('Description', {
                                    label: 'Description',
                                    multiline: true
                                }),
                                PropertyPaneTextField('PrefixText', {
                                    label: 'Button Prefix Text',
                                    description: 'Default: "Subscribe to"'
                                }),
                                PropertyPaneTextField('SuffixText', {
                                    label: 'Button Suffix Text',
                                    description: 'Default: "updates"'
                                })
                            ]
                        },
                        {
                            groupName: 'Font Sizes',
                            groupFields: [
                                PropertyPaneSlider('HeaderFontSize', {
                                    label: 'Header Font Size (px)',
                                    min: 12,
                                    max: 48,
                                    step: 1,
                                    value: this.properties.HeaderFontSize || 22,
                                    showValue: true
                                }),
                                PropertyPaneSlider('DescriptionFontSize', {
                                    label: 'Description Font Size (px)',
                                    min: 10,
                                    max: 24,
                                    step: 1,
                                    value: this.properties.DescriptionFontSize || 14,
                                    showValue: true
                                }),
                                PropertyPaneSlider('ButtonFontSize', {
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
}(BaseClientSideWebPart));
export default SubscribePortalUpdateWebPart;
//vishnu
//# sourceMappingURL=SubscribePortalUpdateWebPart.js.map