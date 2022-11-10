import * as React from 'react';

import { GraphFI } from '@pnp/graph';
import { SPFI } from '@pnp/sp';

type PnpProviderType = {
  graph: GraphFI;
  sp: SPFI;
};

const PnpContext = React.createContext<PnpProviderType | null>(null);

const PnpProvider: React.FC<
  PnpProviderType & { children?: React.ReactNode }
> = ({ children, graph, sp }) => {
  return (
    <PnpContext.Provider value={{ graph, sp }}>{children}</PnpContext.Provider>
  );
};

const usePnpContext = (): PnpProviderType => {
  const pnpContextValue = React.useContext(PnpContext);
  if (pnpContextValue === null)
    throw new Error("Use this function within a GraphProvider");
  return pnpContextValue;
};

export { PnpProvider, PnpProviderType, usePnpContext };
