import React, { forwardRef } from "react";
import SvgSharePoint from "./SharePoint";

export const SharePointIcon = forwardRef<
  SVGSVGElement,
  React.PropsWithChildren<{}>
>((props, ref) => {
  return <SvgSharePoint ref={ref} {...props} />;
});
