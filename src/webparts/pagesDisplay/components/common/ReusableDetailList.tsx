import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  DetailsHeader,
  Selection,
  IDetailsListStyles,
  DetailsRow,
} from "@fluentui/react/lib/DetailsList";
import { mergeStyles } from "@fluentui/react";
import "./styles.css";
import { IColumnInfo } from "../PagesList/PagesService";
import { WebPartContext } from "@microsoft/sp-webpart-base";

const gridStyles: Partial<IDetailsListStyles> = {
  root: {},
  headerWrapper: {},
};

const customHeaderClass = mergeStyles({
  backgroundColor: "#efefef",
  color: "white",
  paddingTop: 0,
  paddingBottom: 0,
  selectors: {
    "& .ms-DetailsHeader": {
      backgroundColor: "#0078d4",
      borderBottom: "1px solid #ccc",
    },
  },
});

export interface IReusableDetailListcomponentsProps {
  columns: (
    columns: IColumnInfo[],
    context: WebPartContext,
    currentUser: any,
    onColumnClick: any,
    sortBy: string,
    isDecending: boolean,
    setShowFilter: (column: IColumn, columnType: string) => void
  ) => IColumn[];
  columnInfos: IColumnInfo[];
  currentUser: any;
  context: WebPartContext;
  setShowFilter: (column: IColumn, columnType: string) => void;
  updateSelection: (selection: Selection) => void;
  items: any[];
  sortPages: (column: IColumn, isAscending: boolean) => void;
  sortBy: string;
  siteUrl: string;
  isDecending: boolean;
  loadMoreItems: () => void; // New prop to load more items
}

export class ReusableDetailList extends React.Component<
  IReusableDetailListcomponentsProps,
  { isLoading: boolean }
> {
  private _selection: Selection;

  private observer: any = null;
  private lastItemRef: any = React.createRef(); // Ref for the last item;
  constructor(components: IReusableDetailListcomponentsProps) {
    super(components);

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.props.updateSelection(this._selection);
      },
      getKey: this._getKey,
    });
  }

  componentDidMount() {
    window.addEventListener("contentLoaded", this.handleContentLoaded);
    this.setupIntersectionObserver();
  }

  componentWillUnmount() {
    window.removeEventListener("contentLoaded", this.handleContentLoaded);
    if (this.observer) {
      this.observer.disconnect();
    }
  }

  componentDidUpdate(prevProps: IReusableDetailListcomponentsProps) {
    window.dispatchEvent(new Event("contentLoaded"));
    this.setupIntersectionObserver();
  }

  setupIntersectionObserver() {
    // Clean up any existing observer
    if (this.observer) {
      this.observer.disconnect();
    }

    // Set up the IntersectionObserver
    this.observer = new IntersectionObserver(this.handleObserver, {
      root: null, // Use the viewport as the root
      rootMargin: "100px", // Load more when the item is 100px from the viewport
      threshold: 0.1, // Trigger when 10% of the item is visible
    });

    // Observe the last item in the list
    if (this.lastItemRef.current) {
      this.observer.observe(this.lastItemRef.current);
    }
  }

  // Function to handle the IntersectionObserver behavior
  handleObserver = (entries: any) => {
    const target = entries[0];
    if (target.isIntersecting) {
      // Trigger loadMoreItems when the last item is visible
      this.props.loadMoreItems();
    }
  };

  handleContentLoaded = () => {
    const navSection: HTMLElement | null =
      document.querySelector(".custom-nav");
    const detailSection: HTMLElement | null =
      document.querySelector(".detail-display");

    function adjustNavHeight() {
      if (navSection && detailSection) {
        const detailHeight = detailSection.offsetHeight;
        const minHeight = 500;
        navSection.style.height = `${Math.max(detailHeight, minHeight)}px`;
      }
    }

    adjustNavHeight();
    window.addEventListener("resize", adjustNavHeight);
  };

  _onRenderDetailsHeader = (components: any) => {
    if (!components) {
      return null;
    }

    return (
      <DetailsHeader
        {...components}
        className="stickyHeader"
        styles={{
          root: customHeaderClass,
        }}
      />
    );
  };

  public render() {
    const {
      columnInfos,
      currentUser,
      context,
      columns,
      items,
      sortPages,
      sortBy,
      isDecending,
      setShowFilter,
    } = this.props;

    return (
      <div style={{ maxHeight: "600px", overflowY: "auto" }}>
        <DetailsList
          styles={gridStyles}
          items={items}
          compact={true}
          columns={columns(
            columnInfos,
            context,
            currentUser,
            sortPages,
            sortBy,
            isDecending,
            setShowFilter
          )}
          selectionMode={SelectionMode.single}
          selection={this._selection}
          getKey={this._getKey}
          setKey="none"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          onRenderDetailsHeader={this._onRenderDetailsHeader}
          onItemInvoked={this._onItemInvoked}
          className="detailList"
          onRenderRow={(props) => {
            if (!props) return null;

            const isLastRow = props.itemIndex === items.length - 1;

            // Attach ref to the last item
            return (
              <div ref={isLastRow ? this.lastItemRef : null}>
                <DetailsRow {...props} />
              </div>
            );
          }}
        />
      </div>
    );
  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }

  private _onItemInvoked = (item: any): void => {
    window.open(`${this.props.siteUrl}${item.FileRef}`, "_blank");
  };
}
