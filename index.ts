import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { AdvancedLookupSearchControl } from "./components/AdvancedLookupSearchControl";

export class AdvancedLookupSearch
    implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
    private _container!: HTMLDivElement;
    private _notifyOutputChanged!: () => void;
    private _selectedValue: ComponentFramework.LookupValue[] | undefined;

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement,
    ): void {
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;
        this._selectedValue = context.parameters.lookupField.raw ?? undefined;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const props = {
            context: context as unknown as ComponentFramework.Context<
                ComponentFramework.PropertyTypes.LookupProperty
            >,
            currentValue: context.parameters.lookupField.raw ?? undefined,
            searchColumns: context.parameters.searchColumns.raw ?? "name",
            displayColumns: context.parameters.displayColumns.raw ?? "",
            placeholderText: context.parameters.placeholderText.raw ?? "Search...",
            maxResults: context.parameters.maxResults.raw ?? 10,
            isDisabled: context.mode.isControlDisabled,
            onChange: (value: ComponentFramework.LookupValue[] | undefined) => {
                this._selectedValue = value;
                this._notifyOutputChanged();
            },
        };

        ReactDOM.render(
            React.createElement(AdvancedLookupSearchControl, props),
            this._container,
        );
    }

    public getOutputs(): IOutputs {
        return {
            lookupField: this._selectedValue,
        };
    }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this._container);
    }
}
