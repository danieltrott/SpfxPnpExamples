import * as React from "react";
import { SearchBox } from "@fluentui/react";
import { escape } from "@microsoft/sp-lodash-subset";

import { IPnpProvidersProps } from "./IPnpProvidersProps";
import styles from "./PnpProviders.module.scss";
import { useMyLists } from "../hooks/useMyLists";

const PnpProviders: React.FC<IPnpProvidersProps> = ({
  isDarkTheme,
  environmentMessage,
  hasTeamsContext,
  userDisplayName,
}) => {
  const { filteredLists: myLists, searchLists } = useMyLists();

  return (
    <section
      className={`${styles.pnpProviders} ${
        hasTeamsContext ? styles.teams : ""
      }`}
    >
      <div className={styles.welcome}>
        <img
          alt=""
          src={
            isDarkTheme
              ? require("../assets/welcome-dark.png")
              : require("../assets/welcome-light.png")
          }
          className={styles.welcomeImage}
        />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>Environment: {environmentMessage}</div>
      </div>
      <div>
        <h3>Available lists</h3>
        <p>
          <SearchBox
            underlined
            placeholder="Search lists"
            onChange={(_, newValue) => {
              searchLists(newValue);
            }}
          />
        </p>
        <ul className={styles.list}>
          {myLists.map((myList) => {
            return <li key={myList.Id}>{myList.Title}</li>;
          })}
        </ul>
      </div>
    </section>
  );
};

export default PnpProviders;
