import * as React from 'react';
import {
  ContextualMenuItemType,
  DirectionalHint,
  IContextualMenuItem,
  IContextualMenuProps,
} from '@fluentui/react/lib/ContextualMenu';
import { DefaultButton } from '@fluentui/react/lib/Button';

const keys: string[] = [
  'newItem',
  'share',
  'mobile',
  'enablePrint',
  'enableMusic',
  'newSub',
  'emailMessage',
  'calendarEvent',
  'disabledNewSub',
  'disabledEmailMessage',
  'disabledCalendarEvent',
  'splitButtonSubMenuLeftDirection',
  'emailMessageLeft',
  'calendarEventLeft',
  'disabledPrimarySplit',
];

export const ContextualMenuCheckmarksExample: React.FunctionComponent = () => {
  const [selection, setSelection] = React.useState<{ [key: string]: boolean }>({});

  const onToggleSelect = React.useCallback(
    (ev?: React.MouseEvent<HTMLButtonElement>, item?: IContextualMenuItem): void => {
      ev && ev.preventDefault();

      if (item) {
        setSelection({ ...selection, [item.key]: selection[item.key] === undefined ? true : !selection[item.key] });
      }
    },
    [selection],
  );

  const menuProps : any = React.useMemo(
    () => ({
      shouldFocusOnMount: true,
      items: [
        { key: keys[0], text: 'New', canCheck: true, isChecked: selection[keys[0]], onClick: onToggleSelect },
        { key: keys[1], text: 'Share', canCheck: true, isChecked: selection[keys[1]], onClick: onToggleSelect },
        { key: keys[2], text: 'Mobile', canCheck: true, isChecked: selection[keys[2]], onClick: onToggleSelect },
      ],
    }),
    [selection, onToggleSelect],
  );

  return <DefaultButton text="Click for ContextualMenu" menuProps={menuProps} />;
};
