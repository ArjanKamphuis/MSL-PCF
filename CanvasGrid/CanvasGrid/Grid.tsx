import * as React from "react";
import { IObjectWithKey, Selection, SelectionMode } from "@fluentui/react/lib/Selection";
import { Stack } from "@fluentui/react/lib/Stack";
import { Sticky, StickyPositionType } from "@fluentui/react/lib/Sticky";
import { IRenderFunction } from "@fluentui/react/lib/Utilities";
import { ConstrainMode, DetailsList, DetailsListLayoutMode, DetailsRow, IColumn, IDetailsHeaderProps, IDetailsRowProps, IDetailsRowStyles } from "@fluentui/react/lib/DetailsList";
import { ScrollablePane, ScrollbarVisibility } from "@fluentui/react/lib/ScrollablePane";
import { Overlay } from "@fluentui/react/lib/Overlay";
import { useConst, useForceUpdate } from "@fluentui/react-hooks";
import { ContextualMenu, DirectionalHint, IContextualMenuItem, IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
import { IconButton } from "@fluentui/react/lib/Button";
import { Link } from "@fluentui/react/lib/Link";

type DataSet = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord & IObjectWithKey;

export type GridProps = {
    width?: number;
    height?: number;
    columns: ComponentFramework.PropertyHelper.DataSetApi.Column[];
    records: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord>;
    sortedRecordIds: string[];
    hasPreviousPage: boolean;
    hasNextPage: boolean;
    currentPage: number;
    sorting: ComponentFramework.PropertyHelper.DataSetApi.SortStatus[];
    filtering: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression;
    resources: ComponentFramework.Resources;
    itemsLoading: boolean;
    isFullScreen: boolean;
    highlightValue: string | null;
    highlightColor: string | null;
    setSelectedRecords: (ids: string[]) => void;
    onNavigate: (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) => void;
    onSort: (name: string, desc: boolean) => void;
    onFilter: (name: string, filtered: boolean) => void;
    loadFirstPage: () => void;
    loadPreviousPage: () => void;
    loadNextPage: () => void;
    onFullScreen: () => void;
};

const stringFormat = (template: string, ...args: string[]): string => {
    for (const k in args) {
        template = template.replace(`{${k}}`, args[k]);
    }
    return template;
}

const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
    if (props && defaultRender) {
        return (
            <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
                {defaultRender({ ...props })}
            </Sticky>
        );
    }
    return null;
};

const onRenderItemColumn = (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, index?: number, column?: IColumn) => {
    return <>{column?.fieldName && item?.getFormattedValue(column.fieldName)}</>;
};

export const Grid = React.memo((props: GridProps) => {
    const {
        width, height,
        columns,
        records,
        sortedRecordIds,
        hasPreviousPage, hasNextPage,
        currentPage,
        sorting, filtering,
        resources,
        itemsLoading,
        isFullScreen,
        highlightValue, highlightColor,
        setSelectedRecords,
        onNavigate,
        onSort, onFilter,
        loadFirstPage, loadPreviousPage, loadNextPage,
        onFullScreen
    } = props;

    const [isComponentLoading, setIsComponentLoading] = React.useState<boolean>(false);
    const [contextualMenuProps, setContextualMenuProps] = React.useState<IContextualMenuProps>();

    const forceUpdate = useForceUpdate();

    const onSelectionChanged = (): void => {
        const items: DataSet[] = selection.getItems() as DataSet[];
        const selected: string[] = selection.getSelectedIndices().map((index: number) => {
            return items[index] && items[index].getRecordId();
        });
        setSelectedRecords(selected);
        forceUpdate();
    };

    // TODO: find a way to use useCallback and useConst together.
    //
    // const onSelectionChanged = React.useCallback((): void => {
    //     const items: DataSet[] = selection.getItems() as DataSet[];
    //     const selected: string[] = selection.getSelectedIndices().map((index: number) => {
    //         return items[index] && items[index].getRecordId();
    //     });
    //     setSelectedRecords(selected);
    //     forceUpdate();
    // }, [forceUpdate, selection, setSelectedRecords]);

    const selection: Selection = useConst(() => {
        return new Selection({ selectionMode: SelectionMode.single, onSelectionChanged });
    });

    const onContextualMenuDismissed = React.useCallback((): void => {
        setContextualMenuProps(undefined);
    }, []);

    const getContextualMenuProps = React.useCallback((column: IColumn, ev: React.MouseEvent<HTMLElement>): IContextualMenuProps => {
        const menuItems: IContextualMenuItem[] = [
            {
                key: 'aToZ',
                text: resources.getString('Label_SortAZ'),
                iconProps: { iconName: 'SortUp' },
                canCheck: true,
                checked: column.isSorted && !column.isSortedDescending,
                disabled: (column.data as ComponentFramework.PropertyHelper.DataSetApi.Column).disableSorting,
                onClick: () => {
                    onSort(column.key, false);
                    setContextualMenuProps(undefined);
                    setIsComponentLoading(true);
                }
            },
            {
                key: 'zToA',
                text: resources.getString('Label_SortZA'),
                iconProps: { iconName: 'SortDown' },
                canCheck: true,
                checked: column.isSorted && column.isSortedDescending,
                disabled: (column.data as ComponentFramework.PropertyHelper.DataSetApi.Column).disableSorting,
                onClick: () => {
                    onSort(column.key, true);
                    setContextualMenuProps(undefined);
                    setIsComponentLoading(true);
                }
            },
            {
                key: 'filter',
                text: resources.getString('Label_DoesNotContainData'),
                iconProps: { iconName: 'Filter' },
                canCheck: true,
                checked: column.isFiltered,
                onClick: () => {
                    onFilter(column.key, column.isFiltered !== true);
                    setContextualMenuProps(undefined);
                    setIsComponentLoading(true);
                }
            }
        ];
        return {
            items: menuItems,
            target: ev.currentTarget as HTMLElement,
            directionalHint: DirectionalHint.bottomLeftEdge,
            gapSpace: 10,
            isBeakVisible: true,
            onDismiss: onContextualMenuDismissed
        };
    }, [onContextualMenuDismissed, onFilter, onSort, resources]);

    const onColumnContextMenu = React.useCallback((column?: IColumn, ev?: React.MouseEvent<HTMLElement>): void => {
        if (column && ev) {
            setContextualMenuProps(getContextualMenuProps(column, ev));
        }
    }, [getContextualMenuProps]);

    const onColumnClick = React.useCallback((ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        if (column && ev) {
            setContextualMenuProps(getContextualMenuProps(column, ev));
        }
    }, [getContextualMenuProps]);

    const onRenderRow: IRenderFunction<IDetailsRowProps> = (props) => {
        const customStyles: Partial<IDetailsRowStyles> = {};
        if (props?.item) {
            const item = props.item as DataSet | undefined;
            if (highlightColor && highlightValue && item?.getValue('HighlightIndicator') == highlightValue) {
                customStyles.root = { backgroundColor: highlightColor };
            }
            return <DetailsRow {...props} styles={customStyles} />
        }
        return null;
    };

    const items: (DataSet | undefined)[] = React.useMemo(() => {
        setIsComponentLoading(false);
        return sortedRecordIds.map(id => records[id]);
    }, [records, sortedRecordIds]);

    const gridColumns: IColumn[] = React.useMemo(() => {
        return columns
            .filter(col => !col.isHidden && col.order >= 0)
            .sort((a, b) => a.order - b.order)
            .map(col => {
                const sortOn: ComponentFramework.PropertyHelper.DataSetApi.SortStatus | undefined = sorting?.find(s => s.name === col.name);
                const filtered: ComponentFramework.PropertyHelper.DataSetApi.ConditionExpression | undefined = filtering?.conditions?.find(f => f.attributeName === col.name);
                return {
                    key: col.name,
                    name: col.displayName,
                    fieldName: col.name,
                    isSorted: sortOn != null,
                    isSortedDescending: sortOn?.sortDirection === 1,
                    isResizable: true,
                    isFiltered: filtered != null,
                    data: col,
                    onColumnContextMenu,
                    onColumnClick
                } as IColumn;
            });
    }, [columns, filtering?.conditions, sorting, onColumnContextMenu, onColumnClick]);

    const rootContainerStyle: React.CSSProperties = React.useMemo(() => {
        return { width, height };
    }, [width, height]);

    const onFirstPage = React.useCallback((): void => {
        setIsComponentLoading(true);
        loadFirstPage();
    }, [loadFirstPage]);
    const onPreviousPage = React.useCallback((): void => {
        setIsComponentLoading(true);
        loadPreviousPage();
    }, [loadPreviousPage]);
    const onNextPage = React.useCallback((): void => {
        setIsComponentLoading(true);
        loadNextPage();
    }, [loadNextPage]);

    return (
        <Stack verticalFill grow style={rootContainerStyle}>
            <Stack.Item grow style={{ position: 'relative', backgroundColor: 'white' }}>
                <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                    <DetailsList
                        columns={gridColumns}
                        onRenderItemColumn={onRenderItemColumn}
                        onRenderDetailsHeader={onRenderDetailsHeader}
                        items={items}
                        setKey={`set${currentPage}`}
                        initialFocusedIndex={0}
                        checkButtonAriaLabel="select row"
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        constrainMode={ConstrainMode.unconstrained}
                        selection={selection}
                        onItemInvoked={onNavigate}
                        onRenderRow={onRenderRow}
                    ></DetailsList>
                    {contextualMenuProps && <ContextualMenu {...contextualMenuProps} />}
                </ScrollablePane>
                {(itemsLoading || isComponentLoading) && <Overlay />}
            </Stack.Item>
            <Stack.Item>
                <Stack horizontal style={{ width: '100%', paddingLeft: 8, paddingRight: 8 }}>
                    <Stack.Item grow align="center">
                        {!isFullScreen && <Link onClick={onFullScreen}>{resources.getString('Label_ShowFullScreen')}</Link>}
                    </Stack.Item>
                    <IconButton alt="First Page" iconProps={{ iconName: 'Rewind' }} disabled={!hasPreviousPage || isComponentLoading || itemsLoading} onClick={onFirstPage} />
                    <IconButton alt="Previous Page" iconProps={{ iconName: 'Previous' }} disabled={!hasPreviousPage || isComponentLoading || itemsLoading} onClick={onPreviousPage} />
                    <Stack.Item align="center">
                        {stringFormat(resources.getString('Label_Grid_Footer'), currentPage.toString(), selection.getSelectedCount().toString())}
                    </Stack.Item>
                    <IconButton alt="Next Page" iconProps={{ iconName: 'Next' }} disabled={!hasNextPage || isComponentLoading || itemsLoading} onClick={onNextPage} />
                </Stack>
            </Stack.Item>
        </Stack>
    );
});

Grid.displayName = 'Grid';
