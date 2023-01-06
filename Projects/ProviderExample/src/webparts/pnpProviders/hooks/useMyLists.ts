import "@pnp/sp/webs";
import "@pnp/sp/lists";

import { useEffect, useState } from "react";

import { Log } from "@microsoft/sp-core-library";
import { debounce } from "@microsoft/sp-lodash-subset";
import { IListInfo } from "@pnp/sp/lists";
import { usePnpContext } from "../provider/PnpProvider";

const LOG_SOURCE = "useMyLists";

type UseListsType = {
  filteredLists: IListInfo[];
  searchLists: (searchTerm?: string) => void;
};

export const useMyLists = (): UseListsType => {
  const [myLists, setMyLists] = useState<IListInfo[]>([]);
  const [filteredLists, setFilteredLists] = useState<IListInfo[]>([]);
  const { sp } = usePnpContext();

  const loadMyLists = async (): Promise<IListInfo[]> => {
    return await sp.web.lists.orderBy("Title")();
  };

  const searchLists = debounce((searchTerm: string = ""): void => {
    const filtered = myLists.filter(
      (list) => list.Title.toUpperCase().indexOf(searchTerm.toUpperCase()) > -1
    );
    setFilteredLists(filtered);
  }, 300);

  useEffect(() => {
    loadMyLists()
      .then((lists) => {
        Log.verbose(LOG_SOURCE, "Loaded lists");
        setMyLists(lists);
      })
      .catch((err) => {
        Log.warn(LOG_SOURCE, "Error loading lists");
        Log.error(LOG_SOURCE, err);
      });
  }, [sp]);

  useEffect(() => {
    searchLists();
  }, [myLists]);

  return { filteredLists, searchLists };
};
