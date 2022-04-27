import { IStackTokens, mergeStyleSets } from "office-ui-fabric-react";

export const stackTokens: IStackTokens = { childrenGap: 20};//, maxWidth:1000 

export const styles = mergeStyleSets({
  checkbox: {
    padding: 5,
  },
  selectAllCheckbox:{
    padding: 5
  },
  button: {
      margin: 10
  }
});