import * as React from 'react';
import styles from './AadGroupPeople.module.scss';
import { IAadGroupPeopleProps } from './IAadGroupPeopleProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export default class AadGroupPeople extends React.Component<IAadGroupPeopleProps, {}> {
  public render(): React.ReactElement<IAadGroupPeopleProps> {
    return (
      <div className={ styles.aadGroupPeople }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>{escape(this.props.groupName[1])}</span>
                {this.props.members.map((user: any) => {
                  return (<div className={styles.description}>
                    <Persona 
                    text={user.displayName}
                    secondaryText={user.userPrincipalName}
                    imageUrl={user.PictureUrl}
                    size={PersonaSize.size48}
                    className={styles.persona}
                    />
                  </div>);
                })}
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
