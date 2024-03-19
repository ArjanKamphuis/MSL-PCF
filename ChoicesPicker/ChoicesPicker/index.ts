import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as React from "react";
import { createRoot, Root as ReactRoot } from 'react-dom/client';
import { initializeIcons } from "@fluentui/react/lib/Icons";
import { ChoicesPickerComponent, ChoicesPickerComponentProps } from "./ChoicesPickerComponent";

initializeIcons(undefined, { disableWarnings: true });

const SmallFormFactorMaxWidth: number = 350;

const enum FormFactors {
    Unknown = 0,
    Desktop = 1,
    Tablet = 2,
    Phone = 3
}

export class ChoicesPicker implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _notifyOutputChanged: () => void;
    private _rootContainer: HTMLDivElement;
    private _reactRoot: ReactRoot;
    private _selectedValue?: number;

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
        this._notifyOutputChanged = notifyOutputChanged;
        this._rootContainer = container;
        this._reactRoot = createRoot(this._rootContainer);
        
        context.mode.trackContainerResize(true);
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        const { value, configuration } = context.parameters;

        let disabled: boolean = context.mode.isControlDisabled;
        let masked: boolean = false;
        if (value.security) {
            disabled = disabled || !value.security.editable;
            masked = !value.security.readable;
        }


        if (value && value.attributes && configuration) {
            const props: ChoicesPickerComponentProps = {
                label: value.attributes.DisplayName,
                options: value.attributes.Options,
                configuration: configuration.raw,
                value: value.raw,
                onChange: this.onChange,
                disabled, masked,
                formFactor: context.client.getFormFactor() == FormFactors.Phone || context.mode.allocatedWidth < SmallFormFactorMaxWidth ? 'small' : 'large'
            };
            this._reactRoot.render(React.createElement(ChoicesPickerComponent, props));
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return { value: this._selectedValue } as IOutputs;
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        this._reactRoot.unmount();
    }

    private onChange = (newValue?: number): void => {
        this._selectedValue = newValue;
        this._notifyOutputChanged();
    }
}
