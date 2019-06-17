import * as React from 'react';
import styles from './SpFxDidp.module.scss';
import { ISpFxDidpProps } from './ISpFxDidpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, autobind } from 'office-ui-fabric-react';

export default class SpFxDidp extends React.Component<ISpFxDidpProps, {}> {
    @autobind
    private _testHandler(): void {
        this.props.sendCommand("command");
        alert("send 'command'");
    }
    public render(): React.ReactElement<ISpFxDidpProps> {
        return (
            <div className={styles.spFxDidp}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <span className={styles.title}>Welcome to SharePoint!</span>
                        </div>
                        <div className={styles.column}>
                            <PrimaryButton text="Send Command" onClick={this._testHandler} />
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}
