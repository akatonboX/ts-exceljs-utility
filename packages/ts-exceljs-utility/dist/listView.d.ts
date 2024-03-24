import React, { CSSProperties } from "react";
export declare const ListView: (props: {
    headers: {
        name: string;
        label: React.ReactNode;
        defaultWidth: number;
    }[];
    hasRowHeader?: boolean | undefined;
    rowHeaderWidth?: number | undefined;
    rowHeight?: number | undefined;
    verticvalLineStyle?: CSSProperties["border"];
    horizontalLineStyle?: CSSProperties["border"];
    resizeColumnEnabled?: boolean | undefined;
    onResizeColumn?: ((name: string, newWidth: number) => void) | undefined;
    children: React.ReactElement<ListViewRowProps, typeof ListViewRow>[] | React.ReactElement<ListViewRowProps, typeof ListViewRow>;
}) => React.JSX.Element;
type ListViewRowProps = {
    header?: React.ReactNode;
    children: React.ReactElement<ListViewItemProps, typeof ListViewItem>[] | React.ReactElement<ListViewItemProps, typeof ListViewItem>;
};
export declare const ListViewRow: (props: ListViewRowProps) => React.JSX.Element;
type ListViewItemProps = {
    name: string;
    children: React.ReactNode;
};
export declare const ListViewItem: (props: ListViewItemProps) => React.JSX.Element;
export {};
