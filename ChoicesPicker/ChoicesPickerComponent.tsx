// use this to reduce size of bundle or alternatively use tree-shaking
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';

export interface ChoicesPickerComponentProps {
    label: string; // Used to label the component. This is bound to the metadata field label that's provided by the parent context, using the UI language selected inside the model-driven app.
    value: number | null; // Linked to the input property defined in the manifest. This can be null when the record is new or the field is not set. TypeScript null is used rather than undefined when passing/returning property values.
    options: ComponentFramework.PropertyHelper.OptionMetadata[]; // When a code component is bound to a choices column in a model-driven app, the property contains the OptionMetadata that describes the choices available. You pass this to the component so it can render each item.
    configuration: string | null; // The purpose of the component is to show an icon for each choice available. The configuration is provided by the app maker when they add the code component to a form. This property accepts a JSON string that maps each numeric choice value to a Fluent UI icon name. For example, {"0":"ContactInfo","1":"Send","2":"Phone"}.
    onChange: (newValue: number | undefined) => void; // When the user changes the choices selection, the React component triggers the onChange event. The code component then calls the notifyOutputChanged so that the model-driven app can update the column with the new value.
    disabled: boolean,
    masked: boolean,
    formFactor: 'small' | 'large' // responsive screen size
}

const iconStyles = { marginRight: '8px' };

const onRenderOption = (option?: IDropdownOption): JSX.Element => {
   if (option) {
       return (
         <div>
             {option.data && option.data.icon && (
               <Icon
                   style={iconStyles}
                   iconName={option.data.icon}
                   aria-hidden="true"
                   title={option.data.icon} />
             )}
             <span>{option.text}</span>
         </div>
       );
   }
   return <></>;
};

const onRenderTitle = (options?: IDropdownOption[]): JSX.Element => {
   if (options) {
       return onRenderOption(options[0]);
   }
   return <></>;
};

export const ChoicesPickerComponent = React.memo((props: ChoicesPickerComponentProps) => {
    const { label, value, options, configuration, disabled, masked, formFactor, onChange } = props;
    const valueKey = value != null ? value.toString() : undefined;
    const items = React.useMemo(() => {
        let iconMapping: Record<number, string> = {};
        let configError: string | undefined;
        if (configuration) {
            try {
                iconMapping = JSON.parse(configuration) as Record<number, string>;
            } catch {
                configError = `Invalid configuration: '${configuration}'`;
            }
        }

        return {
            error: configError,
            choices: options.map((item) => {
                return {
                    key: item.Value.toString(),
                    value: item.Value,
                    text: item.Label,
                    iconProps: { iconName: iconMapping[item.Value] }, // configuration needs to map to {"0":"ContactInfo","1":"Send","2":"Phone"}
                } as IChoiceGroupOption;
            }),
        };
    }, [options, configuration]);

    const onChangeChoiceGroup = React.useCallback(
        (ev?: unknown, option?: IChoiceGroupOption): void => {
            onChange(option ? (option.value as number) : undefined);
        },
        [onChange],
    );

    const onChangeDropDown = React.useCallback(
        (ev: unknown, option?: IDropdownOption): void => {
            onChange(option ? (option.data.value as number) : undefined);
        },
        [onChange],
    );
    
    return (
        <>
            {items.error}
            {masked && '****'}

            {!items.error && !masked && (
                <ChoiceGroup
                    label={label}
                    options={items.choices}
                    selectedKey={valueKey}
                    disabled={disabled}
                    onChange={onChangeChoiceGroup}
                />
            )}
        </>
    );
});
ChoicesPickerComponent.displayName = 'ChoicesPickerComponent';