// https://docs.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference
export { };

type Override<T, U> = Omit<T, keyof U> & U

declare global {
    const Xrm: {
        Navigation: Navigation
        WebApi: WebApi
        Utility: Utility
        /** @deprecated use formContext instead */
        Page: any
    }

    type Utility = {
        showProgressIndicator: (message: string) => void
        closeProgressIndicator: () => void
        getGlobalContext: () => GlobalContext
    }

    type Navigation = {
        openAlertDialog: OpenAlertDialog
        openConfirmDialog: OpenConfirmDialog
        openForm: (entityFormOptions: EntityFormOptions, formParameters: any) => Promise<{ savedEntityReference: { entityType: string, id: string, name: string }[] | null }>
    }

    type EntityFormOptions = {
        entityName: string
        entityId?: string
        formId?: string
        cmdbar?: boolean
        createFromEntity?: { entityType: string, id: string, name?: string }
        height?: number
        width?: number
        navbar?: "on" | "off" | "entity"
        openInNewWindow?: boolean
        windowPosition?: 1 | 2
        relationship?: { attributeName: string, name: string, navigationPropertyName: string, relationshipType: 0 | 1, roleType: 1 | 2 }
        selectedStageId?: string
        useQuickCreateForm?: boolean
    }

    type WebApi = {
        online: any;
        execute: (request: { getMetadata: () => any }) => Promise<{ headers: any, ok: boolean, status: number, statusText: string, url: string, json: Promise<any>, text: Promise<string> }>
        createRecord: (entityLogicalName: string, data: any) => Promise<{ entityType: string, id: string }>
        retrieveRecord: (entityLogicalName: string, id: string, options?: string) => Promise<{ [key: string]: any }>
        retrieveMultipleRecords: (entityLogicalName: string,
            options?: string,
            maxPageSize?: number,
            successCallback?: () => any,
            errorCallback?: () => any) => Promise<{ entities: any[] }>
        updateRecord: (entityLogicalName: string, id: string, data: any) => Promise<{ entityType: string, id: string }>
        deleteRecord: (entityLogicalName: string, id: string) => Promise<{ entityType: string, id: string, name: string }>
    }

    type AlertStrings = { confirmButtonLabel?: string, text: string, title?: string }
    type DialogOptions = { height?: number, width?: number }

    type ConfirmStrings = { cancelButtonLabel?: string, confirmButtonLabel?: string, subtitle?: string, text: string, title?: string }

    type OpenAlertDialog = (alertStrings: AlertStrings, alertOptions?: DialogOptions, closeCallback?: () => any, alertCallback?: () => any) => Promise<void>

    type ConfirmDialogResponse = { confirmed: boolean }
    type OpenConfirmDialog = (confirmStrings: ConfirmStrings, confirmOptions?: DialogOptions, successCallback?: () => any, errorCallback?: () => any) => Promise<ConfirmDialogResponse>

    type UserSettings = {
        roles: {
            get: () => { id: string, name: string }[]
        }
        userId: string
        userName: string
    }

    type GlobalContext = {
        client: { getClient: () => "Web" | "Outlook" | "Mobile" };
        userSettings: UserSettings
    }

    type ExecutionContext<T, I> = {
        getFormContext: () => FormContext<I>
        getEventSource: () => any
        getEventArgs: () => T
    }

    type FormContext<Entity> = {
        getAttribute: <Field extends keyof Entity> (id: Field) => Entity[Field]
        getControl: (id: string) => Control<Entity>
        data: FormContextData
        ui: FormContextUi<Entity>
    }

    type FormContextData = {
        entity: FormContextDataEntity
        process: FormContextDataProcess
        getIsDirty: () => boolean
        isValid: () => boolean
        refresh: (save?: boolean) => Promise<void>
        save: (saveOptions?: { saveMode?: number }) => Promise<void>
    }

    type FormContextDataEntity = {
        getIsDirty: () => boolean
        getDataXml: () => string
        getId: () => string
        addOnSave: (onSave: ExecutionContext<SaveEventArgs, any>) => void
        getEntityName: () => string
        getEntityReference: () => { entityType: string, id: string, name?: string }
        getPrimaryAttributeValue: () => string
        isValid: () => string
        removeOnSave: (onSave: (context: ExecutionContext<SaveEventArgs, any>) => void) => void
        save: (saveOption: "saveandclose" | "saveandnew") => void
    }

    type FormContextDataProcess = {
        addOnPreProcessStatusChange: (onPreChange: (context: ExecutionContext<void, any>) => void) => void
        removeOnPreProcessStatusChange: (onPreChange: (context: ExecutionContext<void, any>) => void) => void
        addOnProcessStatusChange: (onChange: (context: ExecutionContext<void, any>) => void) => void
        removeOnProcessStatusChange: (onChange: (context: ExecutionContext<void, any>) => void) => void
        addOnPreStageChange: (onPreChange: (context: ExecutionContext<void, any>) => void) => void
        removeOnPreStageChange: (onPreChange: (context: ExecutionContext<void, any>) => void) => void
        addOnStageChange: (onChange: (context: ExecutionContext<void, any>) => void) => void
        removeOnStageChange: (onChange: (context: ExecutionContext<void, any>) => void) => void
        addOnStageSelected: (onSelected: (context: ExecutionContext<void, any>) => void) => void
        removeOnStageSelected: (onSelected: (context: ExecutionContext<void, any>) => void) => void

        getActiveProcess: () => Process
        setActiveProcess: (processId: string, callbackFunction: (status: "success" | "invalid") => void) => void
        getEnabledProcesses: (callbackFunction: (processes: any) => void) => void

        moveNext: (callbackFunction: (status: "success" | "crossEntity" | "end" | "invalid" | "dirtyForm" | "stageGate") => void) => void
        movePrevious: (callbackFunction: (status: "success" | "crossEntity" | "beginning" | "invalid" | "dirtyForm" | "stageGate" | "preventDefault") => void) => void
    }

    type Process = {
        getId: () => string
        getName: () => string
        getStages: () => Stage
        isRendered: () => boolean
    }

    type Stage = {
        getCategory: () => { getValue: () => number }
        getEntityName: () => string
        getId: () => string
        getName: () => string
        getNavigationBehavior: () => { allowCreateNew: () => boolean }
        getStatus: () => "active" | "inactive"
        getSteps: () => Step[]
    }

    type Step = {
        getAttribute: () => string
        getName: () => string
        getProgress: () => number
        isRequired: () => boolean
        setProgress: (stepProgress: number, message: string) => "invalid" | "success"
    }

    type FormContextUi<T> = {
        setFormNotification: (message: string, level: "ERROR" | "WARNING" | "INFO", uniqueId: string) => void
        clearFormNotification: (id: string) => void
        close: () => void
        getFormType: () => number
        tabs: { get: (tabId: string) => Tab<T> }
        refreshRibbon: (refreshAll?: boolean) => void
        addOnLoad: (onLoad: ExecutionContext<LoadEventArgs, any>) => void
        removeOnLoad: (onLoad: ExecutionContext<LoadEventArgs, any>) => void
        getViewPortHeight: () => number
        getViewPortWidth: () => number
        setFormEntityName: (name: string) => void
    }

    type TabContentType = "cardSections" | "singleComponent"
    type TabDisplayState = "expanded" | "collapsed"

    type Tab<T> = {
        sections: {
            get: (sectionId: string) => Section<T>
        }
        getContentType: () => TabContentType
        getDisplayState: () => TabDisplayState
        getLabel: () => string
        getName: () => string
        getParent: FormContextUi<T>
        getVisible: () => boolean
        setLabel: (label: string) => void
        setVisible: (visible: boolean) => void
        addTabStateChange: (context: ExecutionContext<void, any>) => void
        removeTabStateChange: (context: ExecutionContext<void, any>) => void
        setContentType: (type: TabContentType) => void
        setDisplayState: (type: TabDisplayState) => void
        setFocus: () => void
    }

    type Section<T> = {
        controls: Control<T>[];
        getLabel: () => string
        getName: () => string
        getParent: () => Tab<T>
        getVisible: () => boolean
        setLabel: (label: string) => void
        setVisible: (visible: boolean) => void
    }

    type Option = {
        text: string
        value: number
    }

    type RequirementLevel = "none" | "required" | "recommended"

    type UserPriviledge = { canRead: boolean, canUpdate: boolean, canCreate: boolean }

    type Attribute = {
        getValue: () => any
        setValue: (value: any) => void
        addOnChange: (onChange: (context: ExecutionContext<void, any>) => void) => void
        controls: Control<any>[]
        setSubmitMode: (mode: "always" | "never" | "dirty") => void
        getPartyList: () => boolean
        getMax: () => number
        getMin: () => number
        getName: () => string
        getRequiredLevel: () => RequirementLevel
        setRequiredLevel: (requirementLevel: RequirementLevel) => void
        getUserPriviledge: () => UserPriviledge
        fireOnChange: () => void
        getIsDirty: () => boolean
    }

    type AttributeLookup = Override<Attribute, {
        getValue: () => [{ id: string }]
    }>

    type AttributeString = Override<Attribute, {
        getValue: () => string
    }>

    type AttributeNumber = Override<Attribute, {
        getValue: () => number
    }>

    type AttributeBoolean = Override<Attribute, {
        getValue: () => boolean
    }>

    type AttributeDate = Override<Attribute, {
        getValue: () => boolean
    }>

    type AttributeOptionSet = Override<Attribute, {
        getValue: () => number
        getOption: () => Option
        getOptions: () => Option[]
    }>

    type Control<T> = {
        getValue: () => string
        /** If set to true, the control will not be editable */
        setDisabled: (disabled: boolean) => void
        getVisible: () => boolean
        /** If set to false, the control will not be shown */
        setVisible: (visible: boolean) => void
        setLabel: (label: string) => void
        setFocus: () => void
        addNotification: (options: XrmNotificationOptions) => void
        clearNotification: (id: string) => void
        getOptions: () => Option[]
        addOption: (option: Option, index?: number) => void
        removeOption: (value: number) => void
        clearOptions: () => void
        addCustomFilter: (filter: string, entityLogicalName?: string) => void
        addPreSearch: (onPreSearch: () => void) => void
        addCustomView: (
            viewId: string,
            entityName: string,
            viewDisplayName: string,
            fetchXml: string,
            layoutXml: string,
            isDefault: boolean) => void
    }

    type XrmNotificationOptions = {
        uniqueId: string
        notificationLevel: "ERROR" | "RECOMMENDATION"
        messages: string[]
        actions?: {
            message: string,
            actions: (() => void)[]
        }[]
    }

    type SaveEventArgs = {
        preventDefault: () => void
        getSaveMode: () => number
        isDefaultPrevented: () => boolean
    }
    type LoadEventArgs = {
        getDataLoadState: () => number
    }
}
