import * as React from "react";

export interface LinearInputComponentProps {
    controlValue: ComponentFramework.PropertyTypes.NumberProperty;
    onChange: (newValue: number) => void;
}

export const LinearInputComponent = React.memo(({ controlValue, onChange }: LinearInputComponentProps) => {
    const [value, setValue] = React.useState<number>(controlValue.raw ?? 0);
    function handleSlider(e: React.ChangeEvent<HTMLInputElement>): void {
        const newValue: number = parseInt(e.target.value);
        setValue(newValue);
        onChange(newValue);
    }
    return (
        <div>
            <input type="range" min="0" max="1000" onChange={handleSlider} value={value} className="linearslider" />
            <label>{value}</label>
        </div>
    );
});
