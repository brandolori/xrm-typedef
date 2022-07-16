// https://docs.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference
export { };
declare global {
    const Xrm: {
        Navigation: Navigation
        WebApi: WebApi
    }

    type Navigation = {
        openAlertDialog: OpenAlertDialog
        openConfirmDialog: OpenConfirmDialog
    }


    type WebApi = {
        createRecord: (entityLogicalName: string, data: any) => Promise<{ entityType: string, id: string }>
        retrieveRecord: (entityLogicalName: string, id: string, options?: string) => Promise<any>
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

    type ExecutionContext<T> = {
        getFormContext: () => FormContext
        getEventSource: () => any
        getEventArgs: () => T
    }

    type FormContext = {
        getAttribute: (id: string) => Attribute
        getControl: (id: string) => Control
        data: FormContextData
        ui: FormContextUi
    }

    type FormContextData = {
        entity: FormContextDataEntity
        getIsDirty: () => boolean
        isValid: () => boolean
        refresh: (save: boolean) => Promise<void>
    }

    type FormContextUi = {
        setFormNotification: (
            message: string,
            level: "ERROR" | "WARNING" | "INFO",
            uniqueId: string) => void
        close: () => void
        getFormType: () => number
        tabs: {
            get: (tabId: string) => Tab
        }
        refreshRibbon: (refreshAll?: boolean) => void
    }

    type TabContentType = "cardSections" | "singleComponent"
    type TabDisplayState = "expanded" | "collapsed"

    type Tab = {
        sections: {
            get: (sectionId: string) => Section
        }
        getContentType: () => TabContentType
        getDisplayState: () => TabDisplayState
        getLabel: () => string
        getName: () => string
        getParent: FormContextUi
        getVisible: () => boolean
        setLabel: (label: string) => void
        setVisible: (visible: boolean) => void
        addTabStateChange: (context: ExecutionContext<void>) => void
        removeTabStateChange: (context: ExecutionContext<void>) => void
        setContentType: (type: TabContentType) => void
        setDisplayState: (type: TabDisplayState) => void
        setFocus: () => void
    }

    type Section = {
        getLabel: () => string
        getName: () => string
        getParent: () => Tab
        getVisible: () => boolean
        setLabel: (label: string) => void
        setVisible: (visible: boolean) => void

    }

    type FormContextDataEntity = {
        getIsDirty: () => boolean
        getDataXml: () => string
        getId: () => string
        getEntityName: () => string
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
        addOnChange: (onChange: (context: ExecutionContext<void>) => void) => void
        controls: Control[]
        setSubmitMode: (mode: "always" | "never" | "dirty") => void
        getPartyList: () => boolean
        getMax: () => number
        getMin: () => number
        getName: () => string
        getOption: () => Option
        getOptions: () => Option[]
        getRequiredLevel: () => RequirementLevel
        setRequiredLevel: (requirementLevel: RequirementLevel) => void
        getUserPriviledge: () => UserPriviledge
        fireOnChange: () => void
    }

    type Control = {
        getValue: () => string
        setDisabled: (enabled: boolean) => void
        getVisible: () => boolean
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

    type FormLoadFunction = (context: ExecutionContext<LoadEventArgs>) => void
    type FormSaveFunction = (context: ExecutionContext<SaveEventArgs>) => void

    type GenericCallback = (context: ExecutionContext<void>) => void

    type ButtonActionPrimaryControl = (context: FormContext) => void
    type ButtonEnablePrimaryControl = (context: FormContext) => boolean


}
