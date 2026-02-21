import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from 'react';
import { InventoryDashboardUI } from "./Dashboard";

export class Dashboard implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private _notifyOutputChanged!: () => void;
    private _context!: ComponentFramework.Context<IInputs>;

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary): void {
        console.log("[PCF] init called (Virtual)", { context, state });
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._context.mode.trackContainerResize(true);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        console.log("[PCF] updateView called (Virtual)", { context });

        const props = {
            context: context as any
        };

        return React.createElement(InventoryDashboardUI, props);
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        console.log("[PCF] destroy called");
    }
}
