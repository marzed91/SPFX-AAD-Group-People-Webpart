import * as React from 'react';
import styles from './AadGroupPeople.module.scss';
import { IAadGroupPeopleProps } from './IAadGroupPeopleProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      'mgt-person': any;
      'mgt-people': any;
      'mgt-people-picker': any;
      'mgt-agenda': any;
      'mgt-tasks': any;
      template: any;
    }
  }
}

export default class AadGroupPeople extends React.Component<IAadGroupPeopleProps, {}> {
  public render(): React.ReactElement<IAadGroupPeopleProps> {
    return (
      <div className={ styles.aadGroupPeople }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <h2 className={ styles.title } role="heading">{escape(this.props.groupName[1])}</h2>
                {this.props.members.map((user: any) => {
                  return (<div className={styles.personaTile}>
                    <mgt-person person-query={user.userPrincipalName} show-name show-email person-card="hover" />
                  </div>);
                })}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
