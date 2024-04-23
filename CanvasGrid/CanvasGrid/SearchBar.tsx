import * as React from "react";
import { Dropdown, IDropdownOption, IDropdownStyles } from "@fluentui/react/lib/Dropdown";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { IStackTokens, Stack } from "@fluentui/react/lib/Stack";

export type SearchBarProps = {
    label: string;
    options: IDropdownOption[];
    search: string;
    onSearchChange: (newValue: string) => void;
    onDropdownChange: (newValue: string) => void;
};

const horizontalGapStackTokens: IStackTokens = { childrenGap: 10, padding: 10 };
const dropDownStyles: Partial<IDropdownStyles> = {
    root: { display: 'flex', alignItems: 'center' },
    dropdown: { width: '11.5rem', marginLeft: '0.25rem' }
};

export const SearchBar = React.memo((props: SearchBarProps) => {
    const { label, options, search, onSearchChange, onDropdownChange } = props;
    return (
        <Stack verticalAlign="center" horizontal tokens={horizontalGapStackTokens}>
            <Dropdown
                label={label}
                options={options}
                onChange={(_ev, option?, _index?) => onDropdownChange(option?.key.toString() ?? '')}
                styles={dropDownStyles}
            />
            <SearchBox placeholder="Search..." value={search} onChange={(_ev?, newValue?) => onSearchChange(newValue?.toLowerCase() ?? '')} />
        </Stack>
    )
});

SearchBar.displayName = 'SearchBar';
