import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'SpFxDidpWebPartStrings';
import SpFxDidp from './components/SpFxDidp';
import { ISpFxDidpProps } from './components/ISpFxDidpProps';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { autobind } from 'office-ui-fabric-react';

export interface ISpFxDidpWebPartProps {
    description: string;
}

export interface ICommand {
    command: string;
}

export default class SpFxDidpWebPart extends BaseClientSideWebPart<ISpFxDidpWebPartProps> implements IDynamicDataCallables {
    private _command: ICommand = undefined;

    protected onInit() {
        return super.onInit()
            .then(() => {
                //register command dynamic property
                this.context.dynamicDataSourceManager.initializeSource(this);

            });
    }

    public render(): void {
        const element: React.ReactElement<ISpFxDidpProps> = React.createElement(
            SpFxDidp,
            {
                description: this.properties.description,
                sendCommand: this._sendCommand
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    @autobind
    private _sendCommand(command: string): void {
        this._command = {
            command: command
        };

        this.context.dynamicDataSourceManager.notifyPropertyChanged("command");
        //this.context.dynamicDataProvider.tryGetSource("")
    }

    public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
        return [
            {
                id: 'command',
                title: 'command'
            }
        ];
    }

    /**
     * Return the current value of the specified dynamic data set
     * @param propertyId ID of the dynamic data set to retrieve the value for
     */
    public getPropertyValue(propertyId: string): ICommand {
        switch (propertyId) {
            case 'command':
                return this._command != null ? this._command : undefined;
        }

        throw new Error('Bad property id');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
