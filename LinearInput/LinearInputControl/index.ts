import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as React from 'react';
import { createRoot, Root as ReactRoot } from 'react-dom/client';
import { LinearInputComponent } from "./LinearInputComponent";

export class LinearInputControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _value: number;
    private _notifyOutputChanged: () => void;
    private labelElement: HTMLLabelElement;
    private inputElement: HTMLInputElement;
    private _container: HTMLDivElement;
    private _reactContainer: HTMLDivElement;
    private _reactRoot: ReactRoot;
    private _refreshData: EventListenerOrEventListenerObject;

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
        this._container = document.createElement('div');
        this._reactContainer = document.createElement('div');
        this._reactRoot = createRoot(this._reactContainer);
        this._notifyOutputChanged = notifyOutputChanged;
        this._refreshData = this.refreshData.bind(this);

        this.inputElement = document.createElement('input');
        this.inputElement.setAttribute('type', 'range');
        this.inputElement.addEventListener('input', this._refreshData);

        this.inputElement.setAttribute('min', '1');
        this.inputElement.setAttribute('max', '1000');
        this.inputElement.setAttribute('class', 'linearslider');
        this.inputElement.setAttribute('id', 'linearrangeinput');

        this.labelElement = document.createElement('label');
        this.labelElement.setAttribute('class', 'linearrangelabel');
        this.labelElement.setAttribute('id', 'lrclabel');

        this._value = context.parameters.controlValue.raw ?? 0;
        this.inputElement.setAttribute('value', context.parameters.controlValue.formatted ?? '0');
        this.labelElement.innerHTML = context.parameters.controlValue.formatted ?? '0';

        this._container.appendChild(this.inputElement);
        this._container.appendChild(this.labelElement);
        container.appendChild(this._container);
        container.appendChild(this._reactContainer);
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        this._value = context.parameters.controlValue.raw ?? this._value;
        this.inputElement.setAttribute('value', context.parameters.controlValue.formatted ?? '');
        this.labelElement.innerHTML = context.parameters.controlValue.formatted ?? '';

        this._reactRoot.render(React.createElement(LinearInputComponent,
            { controlValue: context.parameters.controlValue, onChange: this.refreshDataReact }
        ));
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {
            controlValue: this._value
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        this.inputElement.removeEventListener('input', this.refreshData);
        this._reactRoot.unmount();
    }

    private refreshData(evt : Event): void {
        this._value = this.inputElement.value as any as number;
        this.labelElement.innerHTML = this.inputElement.value;
        this._notifyOutputChanged();
    }

    refreshDataReact = (newValue: number): void => {
        this._value = newValue;
        this._notifyOutputChanged();
    }
}
