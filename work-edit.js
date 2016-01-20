
/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *  work-edit.js
 *
 *  Work item model view for Fusion Work Management solution for SharePoint 2013.
 * 
 *  Jason Barkes, Fusion Alliance
 *  http://www.fusionalliance.com
 *  January 10, 2016
 * 
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

var Fusion = Fusion || {};
var Views = Fusion.Views = Fusion.Views || {};

Fusion.Views.WorkItemView = function (editView, _, $) {

    "use strict";

    var
        utils = Fusion.Utils,
        spUtils = Fusion.Utils.SP,
        models = Fusion.Models,

        isLoading = true,
        isLoaded = false,

        defaults = {
            modelVar: "rc",
            loadingElt: "#contentLoading",
            dateFormat: "MM/DD/YYYY",
            eventSource: "request-edit.js",
            durations: {
                loading: 0,
                content: 0,
                error: 10,
                saveError: 0,
                success: 4
            },
            colors: {
                success: "green",
                error: "red"
            }
        },

        msgs = {
            titles: {
                saveError: "Save Error",
                saveSuccess: "Request Saved",
                generalProgress: "Progress",
                generalError: "Error",
                loadError: "Fusion.Views.WorkItemView.load Error"
            },
            modelError: "An unexpected error occurred in the backbone view 'WorkItemView'.",
            generalError: "An error occurred loading this request.  Please refresh this page in your browser or contact your administrator.",
            saveError: "An error occurred saving this request.  It may have been changed by another user.  Please refresh this page or contact your administrator.  Detail: {0}",
            saveSuccess: "Successfully saved this request.",
            saveProgress: "Saving request {0}...",
            techInfo: "Technical information",
            unsaved: "This request has unsaved changes.",
            seeLog: "See the system event log in the target site for more details",

            // SharePoint message overrides
            overrides: [{
                error: "The specified name is already in use",
                display: "The specified 'Request Title' is already in use. Please specify a different title."
            }, {
                error: "does not match the object's ETag value",
                display: "This request appears to have been changed by another user. Please refresh this page before making changes."
            }]
        },

        fields = { // TODO: are 'ELT's needed here with stickit bindings?
            FileLeafRef: {
                name: "FileLeafRef",
                elt: "FileLeafRef",
                label: "Request Title"
            },
            WorkDescription: {
                name: "WorkDescription",
                elt: "WorkDescription",
                label: "Request Description"
            },
            WorkRequestedDueDate: {
                name: "WorkRequestedDueDate",
                elt: "WorkRequestedDueDate",
                label: "Requested Due Date"
            },
            WorkDueDate: {
                name: "WorkDueDate",
                elt: "WorkDueDate",
                label: "Due Date"
            },
            WorkType: {
                name: "WorkType",
                elt: "WorkType",
                label: "Request Type"
            },
            WorkStatus: {
                name: "WorkStatus",
                elt: "WorkStatus",
                label: "Request Status"
            },
            WorkArea: {
                name: "WorkArea",
                elt: "WorkArea",
                label: "Request Area"
            },
            WorkRequestParentId: {
                name: "WorkRequestParentId",
                elt: "WorkRequestParentId",
                label: "Parent Request"
            },
            WorkRequestedForTitle: {
                name: "WorkRequestedFor.Title",
                elt: "WorkRequestedForTitle",
                label: "Requested For"
            },
            WorkRequestedForId: {
                name: "WorkRequestedForId",
                elt: "WorkRequestedForId",
                label: "Requested For ID (hidden)"
            }
        },

        settings = {},
        choices = {},
        bindings = {},
        model = null,
        view = {},

        init = function (options) {
            settings = _.extend(defaults, options);
            _.templateSettings.variable = settings.modelVar;

            window.onbeforeunload = function () {
                if (model && (_.size(model._changeSet) > 0)) {
                    return msgs.unsaved;
                }
            };

            // Since we are automatically updating the model, we want the model to also hold invalid
            // values, otherwise we might be validating something other than the user entered.
            Backbone.Validation.configure({
                forceUpdate: true,
                labelFormatter: "label"
            });

            // Extend the validation callbacks to work with SharePoint errors
            _.extend(Backbone.Validation.callbacks, {
                valid: function (view, attr, selector) {
                    var elt = view.$('[name=' + attr + ']'),
                        group = elt.closest('.form-group');

                    group.removeClass('has-error');
                    group.find(".ms-formvalidation").remove();
                },

                invalid: function (view, attr, error, selector) {
                    var elt = view.$("[name=" + attr + "]"),
                        group = elt.closest(".form-group"),
                        errElt = $("<span id='Error_" + attr + "' class='form-span ms-formvalidation'>" + error + "</span>");

                    group.addClass("has-error");
                    group.find(".ms-formvalidation").remove();
                    errElt.insertAfter(elt)
                }
            });
        },

        create = function (options) {
            var deferred = $.Deferred();

            init(options);

            spUtils.getListItemEntityType(settings.listName)
                .then(function (entityName) {
                    var RequestItem = models.WorkItemModel.extend({
                        site: settings.webUrl,
                        list: settings.listName,
                        itemType: entityName,
                        validation: {
                            FileLeafRef: {
                                required: true,
                                minLength: 4,
                                maxLength: 128,
                            }
                        },

                        labels: {
                            FileLeafRef: fields.FileLeafRef.label
                        }
                    }),
                    requestItem = new RequestItem({
                        Id: settings.itemId
                    }),

                    RequestView = Backbone.View.extend({
                        el: settings.contentElt,
                        model: requestItem,
                        template: _.template($(settings.templateElt).html()),
                        bindings: getBindings(),

                        initialize: function () {
                            model = this.model;
                            this.listenTo(this.model, "sync", this.render);
                            this.listenTo(this.model, "error", this.onerror);
                            this.listenTo(this.model, "change", this.onchanged);
                            this.listenTo(this.model, "change:FileLeafRef", this.ontitlechanged);

                            Backbone.Validation.bind(this);
                        },

                        render: function (model, resp, options) {
                            this.$el.html(this.template());
                            showContent();
                            this.stickit(this.model, this.bindings);
                            Fusion.Events.trigger("sync", this.model);
                            return this;
                        },

                        saveModel: function () {
                            var waitDialog = SP.UI.ModalDialog.showWaitScreenWithNoClose(msgs.titles.generalProgress, msgs.saveProgress.format(settings.itemId)),
                                reloadPage = this.model._changeSet[fields.FileLeafRef.name] && this.model._changeSet[fields.FileLeafRef.name].length > 0 ? true : false;

                            this.model.save()
                                .done(function () {
                                    // Notify others the model has been successfully saved
                                    Fusion.Events.trigger("saved", this.model);

                                    SP.UI.Status.removeAllStatus();
                                    utils.showMessage(msgs.titles.saveSuccess, msgs.saveSuccess, settings.colors.success, settings.durations.success);

                                    if (reloadPage) { // Reload the page if the Title is changed
                                        location.href = location.href; // TODO: fix this (causes empty request fields sometimes)
                                    }
                                })
                                .fail(function (error) {
                                    SP.UI.Status.removeAllStatus();

                                    var spMessage = spUtils.parseError(error), // Extract the message returned from the SharePoint REST error
                                        mappedMessage = _.filter(msgs.overrides, // Get the message override (if defined)
                                            function (msg) {
                                                return spMessage.match(new RegExp(msg.error, "i"))
                                            }
                                        ),
                                        displayMessage = mappedMessage && mappedMessage.length > 0 ? mappedMessage[0].display : spMessage; // Use the override or the parsed SharePoint message

                                    var logger = new spUtils.CompositeLogger();
                                    logger.push({
                                        location: spUtils.LogLocations.console,
                                        level: spUtils.ConsoleLevels.error,
                                        message: error
                                    });
                                    logger.push({
                                        location: spUtils.LogLocations.console,
                                        level: spUtils.ConsoleLevels.log,
                                        message: msgs.seeLog
                                    });
                                    logger.push({
                                        location: spUtils.LogLocations.statusBar,
                                        color: settings.colors.error,
                                        duration: settings.durations.saveError,
                                        title: msgs.titles.saveError,
                                        message: displayMessage
                                    });
                                    logger.push({
                                        location: spUtils.LogLocations.eventLog,
                                        source: settings.eventSource,
                                        eventType: spUtils.eventTypes.Error,
                                        title: msgs.titles.loadError,
                                        message: error
                                    });
                                    logger.process();
                                })
                                .always(function () {
                                    waitDialog.close();
                                });
                        },

                        onchanged: function (model, options) {
                            if (_.size(this.model._changeSet) > 0 && this.model.isValid(true)) {
                                $("#saveProperties").removeAttr("disabled");
                            } else {
                                $("#saveProperties").attr("disabled", "disabled");
                            }
                            Fusion.Events.trigger("changed", this.model);
                        },

                        ontitlechanged: function (model, value, options) {
                            $.logToConsole("[work-edit.js] Model listener: Title has been changed");
                        },

                        onerror: function (model, resp, options) {
                            $.errorToConsole(msgs.modelError);
                            utils.showMessage(msgs.titles.generalError, msgs.generalError, settings.colors.error, settings.durations.error);
                            Fusion.Events.trigger("error", resp);
                        },

                        events: {
                            "click #saveProperties": function (e) {
                                e.preventDefault();
                                this.saveModel();
                            }
                        },

                        formatDate: function (value, options) {
                            return spUtils.formatDate(value, settings.dateFormat);
                        }
                    }),

                    view = new RequestView();
                    deferred.resolve(requestItem);
                })
                .fail(function (data) {
                    deferred.reject(data);
                });

            return deferred.promise();
        },

        load = function (model) {
            var deferred = $.Deferred();

            if (isLoaded === false) {
                getChoices()
                    .then(function () {
                        return model.fetch({
                            select: "FileLeafRef,WorkRequestedFor/Title,WorkRequestedFor/ID,*",
                            expand: "WorkRequestedFor"
                            // orderby: "ID desc",
                            // top: 5,
                            // skip:10
                        });
                    })
                    .done(function (data) {
                        deferred.resolve(data);
                    })
                    .fail(function (data) {
                        hideContent();
                        var errMessage = spUtils.parseError(data),
                            logger = new spUtils.CompositeLogger();
                        logger.push({
                            location: spUtils.LogLocations.console,
                            level: spUtils.ConsoleLevels.error,
                            message: errMessage
                        });
                        logger.push({
                            location: spUtils.LogLocations.console,
                            level: spUtils.ConsoleLevels.log,
                            message: msgs.seeLog
                        });
                        logger.push({
                            location: spUtils.LogLocations.eventLog,
                            source: settings.eventSource,
                            eventType: spUtils.eventTypes.Error,
                            title: msgs.titles.loadError,
                            message: data
                        });
                        logger.process();

                        // Display the error to the user
                        var errTemplate = _.template($(settings.errTemplateElt).html());
                        $(settings.contentElt).after(errTemplate({
                            errText: errMessage
                        }));
                        attachErrorEvents();

                        deferred.reject(data);
                    })
                    .always(function () {
                        hideLoading();
                    });
            }

            isLoaded = true;

            return deferred.promise();
        },

        /**
         * Backbone Stickit model <-> view bindings
         */
        getBindings = function () {
            var bindings = {
                "#FileLeafRef": {
                    observe: fields.FileLeafRef.name,
                    setOptions: {
                        validate: true
                    }
                },
                ".workfor": {
                    observe: fields.WorkRequestedForTitle.name,
                    setOptions: {
                        validate: true
                    },
                    update: function ($el, val, model, options) {
                        var peoplePicker = $el.spGetPeoplePicker();
                        if (!peoplePicker) return;

                        var userInfo = peoplePicker.GetAllUserInfo()[0];
                        if (typeof userInfo !== "undefined" && userInfo.IsResolved) {
                            spUtils.getUserByAccount(encodeURIComponent(userInfo.Key))
                                .done(function (data) {
                                    model.set(fields.WorkRequestedForId.name, data.Id);
                                });
                        }
                    }
                },
                "#WorkDescription": {
                    observe: fields.WorkDescription.name,
                    setOptions: {
                        validate: true
                    }
                },
                "#WorkDueDate": {
                    observe: fields.WorkDueDate.name,
                    setOptions: {
                        validate: true
                    },
                    onGet: "formatDate",
                    onSet: "formatDate"
                },
                "#WorkRequestedDueDate": {
                    observe: fields.WorkRequestedDueDate.name,
                    setOptions: {
                        validate: true
                    },
                    onGet: "formatDate",
                    onSet: "formatDate"
                },
                ".worktype": {
                    observe: fields.WorkType.name,
                    selectOptions: {
                        collection: function () {
                            return choices.WorkType;
                        },
                        labelPath: "name",
                        valuePath: "value",
                        defaultOption: {
                            label: "Select...",
                            value: null
                        }
                    }
                },
                ".workstatus": {
                    observe: fields.WorkStatus.name,
                    selectOptions: {
                        collection: function () {
                            return choices.WorkStatus;
                        },
                        labelPath: "name",
                        valuePath: "value",
                        defaultOption: {
                            label: "Select...",
                            value: null
                        }
                    }
                },
                ".workarea": {
                    observe: fields.WorkArea.name,
                    selectOptions: {
                        collection: function () {
                            return choices.WorkArea;
                        },
                        labelPath: "name",
                        valuePath: "value",
                        defaultOption: {
                            label: "Select...",
                            value: null
                        }
                    }
                },
                ".workparent": {
                    observe: fields.WorkRequestParentId.name,
                    selectOptions: {
                        collection: function () {
                            return choices.FileLeafRef;
                        },
                        labelPath: "name",
                        valuePath: "value",
                        defaultOption: {
                            label: "Select...",
                            value: null
                        }
                    }
                }
            };

            return bindings;
        },

        getChoices = function () {
            var deferred = $.Deferred();

            var promises = [];
            $.each(settings.dropdowns, function (index) {
                var field = this.field;

                if (this.source === "choice") {
                    promises.push(spUtils.getFieldChoices(settings.listName, field));
                } else if (this.source === "list") {
                    promises.push(spUtils.getFieldValues(settings.listName, field, this.exclude));
                }

                promises[index].done(function (data) {
                    choices[field] = [];
                    $.each(data, function (dataIndex) {
                        choices[field].push({
                            name: data[dataIndex],
                            value: data[dataIndex]
                        });
                    });
                });
            });

            $.when.apply($, promises).done(function (data) {
                deferred.resolve(data);
            });

            $.when.apply($, promises).fail(function (data) {
                deferred.reject(data);
            });

            return deferred.promise();
        },

        showLoading = function (duration) {
            $(settings.loadingElt).show((duration) ? duration : settings.durations.loading);
            isLoading = true;
        },

        hideLoading = function (duration) {
            $(settings.loadingElt).hide((duration) ? duration : settings.durations.loading);
            isLoading = false;
        },

        showContent = function (duration) {
            $(settings.contentContainerElt).show((duration) ? duration : settings.durations.content);
        },

        hideContent = function (duration) {
            $(settings.contentContainerElt).hide((duration) ? duration : settings.durations.content);
        },

        attachErrorEvents = function () {
            utils.attachOnce("expandError", "click", function (e) {
                e.preventDefault();

                if ($("#errDetails").is(":visible")) {
                    $("#errDetails").hide();
                    $("#expandError").text(msgs.techInfo);
                } else {
                    $("#errDetails").show();
                    $("#expandError").text(msgs.techInfo);
                }
            });
        },

        //
        // Public API ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        //
        api = $.extend(editView, {
            create: create,
            load: load,
            showLoading: showLoading,
            hideLoading: hideLoading,
            showContent: showContent,
            hideContent: hideContent,
            isLoading: function () {
                return isLoading;
            },
            getDropdownChoices: function () {
                return choices;
            },
            hasChanged: function () {
                return model && (_.size(model._changeSet) > 0);
            },
            getModel: function () {
                return model;
            }
        });

    return api;

}(Fusion.Views.WorkItemView || {}, _, jQuery);
