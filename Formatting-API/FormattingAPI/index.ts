import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class FormattingAPI implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _controlViewRendered: boolean;

    private _decimal: number;
    private _precision: number;
    private _date: Date;

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this._context = context;
        this._controlViewRendered = false;
        this._container = document.createElement('div');
        this._container.classList.add('TSFormatting_Container');
        container.appendChild(this._container);

        this._decimal = context.parameters.InputDecimal.raw ?? 0;
        this._precision = context.parameters.Precision.raw ?? 2;
        this._date = context.parameters.InputDate.raw ?? new Date();
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        if (!this._controlViewRendered) {
            const table: HTMLTableElement = this.createHTMLTableElement();
            this._container.appendChild(table);
            this._controlViewRendered = true;
        }

        let rerender: boolean = false;
        if (this._decimal != context.parameters.InputDecimal.raw) {
            this._decimal = context.parameters.InputDecimal.raw ?? 0;
            rerender = true;
        }
        if (this._precision != context.parameters.Precision.raw) {
            this._precision = context.parameters.Precision.raw ?? 2;
            if (this._precision < 0) this._precision = 0;
            rerender = true;
        }
        if (this._date != context.parameters.InputDate.raw) {
            this._date = context.parameters.InputDate.raw ?? new Date();
            rerender = true;
        }

        if (rerender) {
            let table: HTMLTableElement | null = this._container.querySelector('table');
            if (table) this._container.removeChild(table);
            this._container.appendChild(this.createHTMLTableElement());
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }

    /**
    * Helper method to create an HTML Table Row Element
    * @param key : string value to show in left column cell
    * @param value : string value to show in right column cell
    * @param isHeaderRow : true if method should generate a header row
    */
    private createHTMLTableRowElement(key: string, value: string, isHeaderRow: boolean): HTMLTableRowElement {
        const keyCell: HTMLTableCellElement = this.createHTMLTableCellElement(key, 'FormattingControlSampleHtmlTable_HtmlCell_Key', isHeaderRow);
        const keyValue: HTMLTableCellElement = this.createHTMLTableCellElement(value, 'FormattingControlSampleHtmlTable_HtmlCell_Value', isHeaderRow);

        const rowElement: HTMLTableRowElement = document.createElement('tr');
        rowElement.setAttribute('class', 'FormattingControlSampleHtmlTable_HtmlRow');
        rowElement.appendChild(keyCell);
        rowElement.appendChild(keyValue);

        return rowElement;
    }

    /**
    * Helper method to create an HTML Table Cell Element
    * @param cellValue : string value to inject in the cell
    * @param className : class name for the cell
    * @param isHeaderRow : true if method should generate a header row cell
    */
    private createHTMLTableCellElement(cellValue: string, className: string, isHeaderRow: boolean): HTMLTableCellElement {
        let cellElement: HTMLTableCellElement;

        if (isHeaderRow) {
            cellElement = document.createElement('th');
            cellElement.setAttribute('class', `FormattingControlSampleHtmlTable_HtmlHeaderCell ${className}`);
        } else {
            cellElement = document.createElement('td');
            cellElement.setAttribute('class', `FormattingControlSampleHtmlTable_HtmlCell ${className}`);
        }

        cellElement.appendChild(document.createTextNode(cellValue));
        return cellElement;
    }

    /**
    * Creates an HTML Table that showcases examples of basic methods available to the custom control
    * The left column of the table shows the method name or property that is being used
    * The right column of the table shows the result of that method name or property
    */
    private createHTMLTableElement(): HTMLTableElement {
        const tableElement: HTMLTableElement = document.createElement('table');
        tableElement.setAttribute('class', 'FormattingControlSampleHtmlTable_HtmlTable');

        let key: string = 'Example Method';
        let value: string = 'Result';
        tableElement.appendChild(this.createHTMLTableRowElement(key, value, true));

        key = 'formatCurrency()';
        value = this._context.formatting.formatCurrency(this._decimal, this._precision, "€");
        tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

        key = 'formatDecimal()';
        value = this._context.formatting.formatDecimal(this._decimal, this._precision);
        tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

        key = 'formatInteger()';
        value = this._context.formatting.formatInteger(this._decimal);
        tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

        key = 'formatLanguage()';
        value = this._context.formatting.formatLanguage(1033);
        tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

        key = 'formatDateLong()';
        value = this._context.formatting.formatDateLong(this._date);
        tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

        key = 'getWeekOfYear()';
        value = this._context.formatting.getWeekOfYear(this._date).toString();
        tableElement.appendChild(this.createHTMLTableRowElement(key, value, false));

        return tableElement;
    }
}
