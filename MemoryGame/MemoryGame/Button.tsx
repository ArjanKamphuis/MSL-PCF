import * as React from "react";

type ButtonStyle = 'default' | 'alternative' | 'dark' | 'light' | 'green' | 'red' | 'yellow' | 'purple';
type ButtonClickEventHandler = React.MouseEventHandler<HTMLButtonElement>;

export type ButtonProps = {
    children: React.ReactNode;
    style?: ButtonStyle;
    onClick?: ButtonClickEventHandler;
    disabled?: boolean;
    id?: string;
    className?: string;
    type?: 'submit' | 'reset' | 'button';
};

export const Button = React.memo(({ children, style, onClick, disabled, id, className, type }: ButtonProps) => {
    const classes: string = buttonMap.get(style ?? 'default') + (className ? ` ${className} ` : '');
    return <button className={classes} onClick={onClick} disabled={disabled} id={id} type={type}>{children}</button>;
});
Button.displayName = 'Button';

const buttonMap: Map<ButtonStyle, string> = new Map();
buttonMap.set('default', 'text-white bg-blue-700 hover:bg-blue-800 font-medium rounded-full text-sm px-5 py-2.5 text-center disabled:opacity-50');
buttonMap.set('alternative', 'py-2.5 px-5 text-sm font-medium text-gray-900 bg-white rounded-full border border-gray-200 hover:bg-gray-100 hover:text-blue-700 disabled:opacity-50');
buttonMap.set('dark', 'text-white bg-gray-800 hover:bg-gray-900 font-medium rounded-full text-sm px-5 py-2.5 disabled:opacity-50');
buttonMap.set('light', 'text-gray-900 bg-white border border-gray-300 font-medium rounded-full text-sm px-5 py-2.5 disabled:opacity-50');
buttonMap.set('green', 'text-white bg-green-700 hover:bg-green-800 font-medium rounded-full text-sm px-5 py-2.5 text-center disabled:opacity-50');
buttonMap.set('red', 'text-white bg-red-700 hover:bg-red-800 font-medium rounded-full text-sm px-5 py-2.5 text-center disabled:opacity-50');
buttonMap.set('yellow', 'text-white bg-yellow-400 hover:bg-yellow-500 font-medium rounded-full text-sm px-5 py-2.5 text-center disabled:opacity-50');
buttonMap.set('purple', 'text-white bg-purple-700 hover:bg-purple-800 font-medium rounded-full text-sm px-5 py-2.5 text-center disabled:opacity-50');
