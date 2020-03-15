import * as React from 'react';
import styles from './AadGroupPeople.module.scss';
import { IAadGroupPeopleProps } from './IAadGroupPeopleProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export default class AadGroupPeople extends React.Component<IAadGroupPeopleProps, {}> {
  public render(): React.ReactElement<IAadGroupPeopleProps> {
    console.log("members: " + this.props.members.length);
    return (
      <div className={ styles.aadGroupPeople }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.groupName)}</p>
              <ul>
                {this.props.members.map((value) => {
                  return <li>{value}</li>
                })}
              </ul>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
