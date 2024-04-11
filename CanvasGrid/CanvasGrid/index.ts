import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { createRoot, Root as ReactRoot } from "react-dom/client"
import { Grid, GridProps } from "./Grid";
import { initializeIcons } from "@fluentui/react";

initializeIcons(undefined, { disableWarnings: true });

export class CanvasGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _container: HTMLDivElement;
    private _resources: ComponentFramework.Resources;
    private _isTestHarness: boolean;
    private _reactRoot: ReactRoot;

    private _sortedRecordIds: string[] = [];
    private _records: { [id: string]: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord };
    private _currentPage: number = 1;
    private _filteredRecordCount?: number;
    private _isFullScreen: boolean = false;

    private _setSelectedRecords = (ids: string[]): void => {
        this._context.parameters.records.setSelectedRecordIds(ids);
    };

    private _onNavigate = (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord): void => {
        if (item) {
            this._context.parameters.records.openDatasetItem(item.getNamedReference());
        }
    };

    private _onSort = (name: string, desc: boolean): void => {
        this._context.parameters.records.sorting = [];
        this._context.parameters.records.sorting.push({ name, sortDirection: desc ? 1 : 0 });
        this._context.parameters.records.refresh();
    };

    private _onFilter = (name: string, filter: boolean): void => {
        const filtering: ComponentFramework.PropertyHelper.DataSetApi.Filtering = this._context.parameters.records.filtering;
        if (filter) {
            filtering.setFilter({
                conditions: [{ attributeName: name, conditionOperator: 12 }]
            } as ComponentFramework.PropertyHelper.DataSetApi.FilterExpression);
        } else {
            filtering.clearFilter();
        }
        this._context.parameters.records.refresh();
    };

    private _loadFirstPage = (): void => {
        this._currentPage = 1;
        this._context.parameters.records.paging.loadExactPage(1);
    };
    private _loadPreviousPage = (): void => {
        this._currentPage--;
        this._context.parameters.records.paging.loadExactPage(this._currentPage);
    };
    private _loadNextPage = (): void => {
        this._currentPage++;
        this._context.parameters.records.paging.loadExactPage(this._currentPage);
    };

    private _onFullScreen = (): void => {
        this._context.mode.setFullScreen(true);
    };

    /**
     * Empty constructor.
     */
    constructor() {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._context.mode.trackContainerResize(true);
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;
        this._resources = this._context.resources;
        this._isTestHarness = document.getElementById('control-dimensions') !== null;
        this._reactRoot = createRoot(this._container);
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const dataset: ComponentFramework.PropertyTypes.DataSet = context.parameters.records;
        const paging: ComponentFramework.PropertyHelper.DataSetApi.Paging = context.parameters.records.paging;
        const datasetChanged: boolean = context.updatedProperties.includes('dataset');
        const resetPaging: boolean = datasetChanged && !dataset.loading && !dataset.paging.hasPreviousPage && this._currentPage !== 1;

        if (context.updatedProperties.includes('fullscreen_close')) this._isFullScreen = false;
        if (context.updatedProperties.includes('fullscreen_open')) this._isFullScreen = true;

        if (resetPaging) {
            this._currentPage = 1;
        }
        if (resetPaging || datasetChanged || this._isTestHarness || !this._records) {
            this._records = dataset.records;
            this._sortedRecordIds = dataset.sortedRecordIds;
        }
        if (this._filteredRecordCount !== this._sortedRecordIds.length) {
            this._filteredRecordCount = this._sortedRecordIds.length;
            this._notifyOutputChanged();
        }

        const props: GridProps = {
            width: parseInt(context.mode.allocatedWidth as unknown as string),
            height: parseInt(context.mode.allocatedHeight as unknown as string),
            columns: dataset.columns,
            records: this._records,
            sortedRecordIds: this._sortedRecordIds,
            hasPreviousPage: paging.hasPreviousPage,
            hasNextPage: paging.hasNextPage,
            currentPage: this._currentPage,
            sorting: dataset.sorting,
            filtering: dataset.filtering && dataset.filtering.getFilter(),
            resources: this._resources,
            itemsLoading: dataset.loading,
            isFullScreen: this._isFullScreen,
            highlightValue: this._context.parameters.HighlightValue.raw,
            highlightColor: this._context.parameters.HighlightColor.raw,
            setSelectedRecords: this._setSelectedRecords,
            onNavigate: this._onNavigate,
            onSort: this._onSort,
            onFilter: this._onFilter,
            loadFirstPage: this._loadFirstPage,
            loadPreviousPage: this._loadPreviousPage,
            loadNextPage: this._loadNextPage,
            onFullScreen: this._onFullScreen
        };
        this._reactRoot.render(React.createElement(Grid, props));
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return { FilteredRecordCount: this._filteredRecordCount };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        this._reactRoot.unmount();
    }
}
