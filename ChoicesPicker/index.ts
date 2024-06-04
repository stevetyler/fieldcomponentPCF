import { IInputs, IOutputs } from './generated/ManifestTypes';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react/lib/Icons'; // The import of initializeIcons is required because you're using the Fluent UI icon set. You need to call initializeIcons to load the icons inside the test harness. Inside model-driven apps, they're already initialized.
import { ChoicesPickerComponent } from './ChoicesPickerComponent';

initializeIcons(undefined, { disableWarnings: true });

export class ChoicesPicker implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	notifyOutputChanged: () => void; // Holds a reference to the method used to notify the model-driven app that a user has changed a choice value and the code component is ready to pass it back to the parent context.
    rootContainer: HTMLDivElement; // HTML DOM element that's created to hold the code component inside the model-driven app.
    selectedValue: number | undefined; // Holds the state of the choice selected by the user so that it can be returned inside the getOutputs method.
    context: ComponentFramework.Context<IInputs>; // Power Apps component framework context that's used to read the properties defined in the manifest and other runtime properties, and access API methods such as trackContainerResize.

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
	public init(
		context: ComponentFramework.Context<IInputs>, 
		notifyOutputChanged: () => void, 
		state: ComponentFramework.Dictionary, 
		container: HTMLDivElement): 
		void {
		  this.notifyOutputChanged = notifyOutputChanged;
		  this.rootContainer = container;
		  this.context = context;
	}

	onChange = (newValue: number | undefined): void => {
		this.selectedValue = newValue;
		this.notifyOutputChanged();
    };

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	// Notice that you're pulling the label and options from context.parameters.value, and the value.raw provides the numeric choice selected or null if no value is selected.
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		const { value } = context.parameters;
		const configObject = {"0":"ContactInfo","1":"Send","2":"Phone"};
		
		// must get masked and disabled flags so we don't update them by mistake 
		let disabled = context.mode.isControlDisabled;
		let masked = false;
		if (value.security) {
			disabled = disabled || !value.security.editable;
			masked = !value.security.readable;
		}

		if (value && value.attributes && configObject) {
			ReactDOM.render(
				React.createElement(ChoicesPickerComponent, {
					label: value.attributes.DisplayName,
					options: value.attributes.Options,
					configuration: JSON.stringify(configObject),
					value: value.raw,
					onChange: this.onChange,
					disabled: disabled,
					masked: masked
				}),
				this.rootContainer,
			);
		}
	}

	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return { value: this.selectedValue } as IOutputs;
	}

	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this.rootContainer);
	}
}
