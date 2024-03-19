import * as React from 'react';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';

export interface ChoicesPickerComponentProps {
    label: string;
    value: number | null;
    options: ComponentFramework.PropertyHelper.OptionMetadata[];
    configuration: string | null;
    onChange: (newValue?: number) => void;
    disabled: boolean;
    masked: boolean;
    formFactor: 'small' | 'large';
}

const IconStyles = { marginRight: '8px' };
const onRenderOption = (option?: IDropdownOption): React.JSX.Element => {
    if (!option) return <></>;
    return (
        <div>
            {option.data && option.data.icon && (
                <Icon style={IconStyles} iconName={option.data.icon} aria-hidden title={option.data.icon} />
            )}
            <span>{option.text}</span>
        </div>
    );
}
const onRenderTitle = (options?: IDropdownOption[]): React.JSX.Element => {
    return options ? onRenderOption(options[0]) : <></>;
}

export const ChoicesPickerComponent = React.memo((props: ChoicesPickerComponentProps) => {
    const { label, value, options, configuration, onChange, disabled, masked, formFactor } = props;
    const valueKey: string | undefined = value !== null ? value.toString() : undefined;

    const items = React.useMemo(() => {
        let iconMapping: Record<number, string> = {};
        let configError: string | undefined;
        if (configuration) {
            try {
                iconMapping = JSON.parse(configuration) as Record<number, string>;
            } catch {
                configError = `Invalid configuration '${configuration}'`;
            }
        }

        return {
            error: configError,
            choices: options.map(item => {
                return {
                    key: item.Value.toString(),
                    value: item.Value,
                    text: item.Label,
                    iconProps: { iconName: iconMapping[item.Value] }
                } as IChoiceGroupOption;
            }),
            options: options.map(item => {
                return {
                    key: item.Value.toString(),
                    data: { value: item.Value, icon: iconMapping[item.Value] },
                    text: item.Label
                } as IDropdownOption;
            })
        };
    }, [options, configuration]);

    const onChangeChoiceGroup = React.useCallback(
        (ev?: unknown, option?: IChoiceGroupOption): void => {
            onChange(option ? (option.value as number) : undefined);
        },
        [onChange]
    );

    const onChangeDropDown = React.useCallback(
        (ev?: unknown, option?: IDropdownOption): void => {
            onChange(option ? (option.data.value as number) : undefined);
        },
        [onChange]
    );

    return (
        <>
            {items.error}
            {masked && '****'}
            {formFactor === 'large' && !items.error && !masked && (
                <ChoiceGroup
                    label={label}
                    options={items.choices}
                    selectedKey={valueKey}
                    disabled={disabled}
                    onChange={onChangeChoiceGroup}
                />
            )}
            {formFactor === 'small' && !items.error && !masked && (
                <Dropdown
                    placeholder={'---'}
                    label={label}
                    aria-label={label}
                    options={items.options}
                    selectedKey={valueKey}
                    disabled={disabled}
                    onRenderOption={onRenderOption}
                    onRenderTitle={onRenderTitle}
                    onChange={onChangeDropDown}
                />
            )}
        </>
    );
});
ChoicesPickerComponent.displayName = 'ChoicesPickerComponent';
