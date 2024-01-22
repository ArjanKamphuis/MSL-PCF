import {IInputs, IOutputs} from "./generated/ManifestTypes";

interface Popup extends ComponentFramework.FactoryApi.Popup.Popup {
    popupStyle: object;
    shadowStyle: object;
}

export class PopupComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _container: HTMLDivElement;
    private _popUpService: ComponentFramework.FactoryApi.Popup.PopupService;
    private static _popupName: string = 'loaderPopup';

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
        const popUpContainer: HTMLDivElement = document.createElement('div');
        popUpContainer.setAttribute('class', 'loader-container');
        const popUpContent: HTMLDivElement = document.createElement('div');
        popUpContent.setAttribute('class', 'loader');
        popUpContainer.appendChild(popUpContent);

        const popUpOptions: Popup = {
            closeOnOutsideClick: true,
            content: popUpContainer,
            name: PopupComponent._popupName,
            type: 1,
            popupStyle: {
                'width': '100%',
                'height': '100%',
                'overflow': 'hidden',
                'backgroundColor': 'transparent',
                'display': 'flex',
                'flexDirection': 'column',
                'position': 'absolute',
                //'margin-top': '28px'
            },
            shadowStyle: {
                position: 'absolute',
                width: '100%',
                height: '100%'
            }
        }

        this._popUpService = context.factory.getPopupService();
        this._popUpService.createPopup(popUpOptions);
        
        container.appendChild(this._container);
        this._popUpService.openPopup('loaderPopup');
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Add code to update control view
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
        this._popUpService.closePopup(PopupComponent._popupName);
        this._popUpService.deletePopup(PopupComponent._popupName);
    }
}
