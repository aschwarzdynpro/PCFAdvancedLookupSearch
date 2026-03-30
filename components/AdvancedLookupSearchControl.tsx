import * as React from "react";
import {
    Callout,
    DirectionalHint,
    FocusTrapZone,
    Icon,
    ISearchBoxStyles,
    List,
    MessageBar,
    MessageBarType,
    SearchBox,
    Spinner,
    SpinnerSize,
    Stack,
    Text,
    mergeStyleSets,
} from "@fluentui/react";

export interface SearchResult {
    id: string;
    entityName: string;
    /** Values for the configured display columns, keyed by column logical name */
    columns: Record<string, string>;
    /** Primary display string (first non-empty display column value) */
    displayName: string;
}

export interface AdvancedLookupSearchControlProps {
    /** PCF context, used to call the Dataverse WebAPI */
    context: ComponentFramework.Context<ComponentFramework.PropertyTypes.LookupProperty>;
    /** Currently selected lookup value (may be undefined / empty) */
    currentValue: ComponentFramework.LookupValue[] | undefined;
    /** Semicolon-separated column names to run the search filter against */
    searchColumns: string;
    /**
     * Semicolon-separated column names to show in each result row.
     * Falls back to searchColumns when empty.
     */
    displayColumns: string;
    /** Placeholder shown in the search input */
    placeholderText: string;
    /** Maximum number of records to retrieve per search */
    maxResults: number;
    /** Whether the control is in a read-only / disabled state */
    isDisabled: boolean;
    /** Callback invoked whenever the lookup selection changes */
    onChange: (value: ComponentFramework.LookupValue[] | undefined) => void;
}

const classNames = mergeStyleSets({
    root: {
        position: "relative",
        width: "100%",
    },
    selectedValueRow: {
        display: "flex",
        alignItems: "center",
        gap: 6,
        minHeight: 30,
        padding: "2px 0",
    },
    selectedValueText: {
        flex: 1,
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap",
        fontSize: 14,
        cursor: "default",
    },
    iconButton: {
        cursor: "pointer",
        color: "#605e5c",
        selectors: {
            ":hover": { color: "#323130" },
        },
    },
    callout: {
        maxHeight: 320,
        overflowY: "auto",
        padding: "4px 0",
        minWidth: 200,
    },
    resultItem: {
        padding: "6px 12px",
        cursor: "pointer",
        selectors: {
            ":hover, :focus": {
                backgroundColor: "#f3f2f1",
                outline: "none",
            },
        },
    },
    resultItemFocused: {
        backgroundColor: "#edebe9",
    },
    resultPrimaryText: {
        fontSize: 14,
        fontWeight: 600,
        display: "block",
    },
    resultSecondaryText: {
        fontSize: 12,
        color: "#605e5c",
        display: "block",
    },
    noResults: {
        padding: "8px 12px",
        color: "#605e5c",
        fontSize: 13,
    },
});

const searchBoxStyles: Partial<ISearchBoxStyles> = {
    root: { width: "100%" },
};

/** Debounce delay in milliseconds before a search is triggered */
const DEBOUNCE_MS = 300;

/**
 * Returns the primary-key field name for a Dataverse entity.
 * For most entities the pattern is `{entitylogicalname}id`.
 * Activity entities share the `activityid` key as a fallback.
 */
function getPrimaryKey(entityName: string): string {
    return `${entityName}id`;
}

/**
 * Builds the OData $filter expression that performs a case-insensitive
 * substring match on each of the supplied columns (OR-joined).
 */
function buildFilter(columns: string[], searchText: string): string {
    const escaped = searchText.replace(/'/g, "''");
    return columns
        .map((col) => `contains(${col},'${escaped}')`)
        .join(" or ");
}

export const AdvancedLookupSearchControl: React.FC<AdvancedLookupSearchControlProps> = ({
    context,
    currentValue,
    searchColumns,
    displayColumns,
    placeholderText,
    maxResults,
    isDisabled,
    onChange,
}) => {
    const [isSearching, setIsSearching] = React.useState(false);
    const [searchText, setSearchText] = React.useState("");
    const [results, setResults] = React.useState<SearchResult[]>([]);
    const [isLoading, setIsLoading] = React.useState(false);
    const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
    const [focusedIndex, setFocusedIndex] = React.useState(-1);

    const searchBoxContainerRef = React.useRef<HTMLDivElement>(null);
    const debounceTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

    const isCalloutOpen = isSearching && (isLoading || results.length > 0 || errorMessage !== null);

    const parsedSearchColumns = React.useMemo(
        () => searchColumns.split(";").map((c) => c.trim()).filter(Boolean),
        [searchColumns],
    );

    const parsedDisplayColumns = React.useMemo(() => {
        const cols = displayColumns
            .split(";")
            .map((c) => c.trim())
            .filter(Boolean);
        return cols.length > 0 ? cols : parsedSearchColumns;
    }, [displayColumns, parsedSearchColumns]);

    const currentDisplayName =
        currentValue && currentValue.length > 0 ? currentValue[0].name ?? "" : "";

    // Derive the target entity type from the bound lookup property
    const entityName: string = (
        context.parameters as unknown as { lookupField: ComponentFramework.PropertyTypes.LookupProperty }
    ).lookupField.getTargetEntityType();

    const runSearch = React.useCallback(
        async (text: string) => {
            if (!text || text.trim().length === 0) {
                setResults([]);
                setErrorMessage(null);
                return;
            }

            if (parsedSearchColumns.length === 0) {
                setErrorMessage("No search columns configured.");
                return;
            }

            setIsLoading(true);
            setErrorMessage(null);

            try {
                const allColumns = [
                    ...new Set([...parsedSearchColumns, ...parsedDisplayColumns]),
                ];
                const select = allColumns.join(",");
                const filter = buildFilter(parsedSearchColumns, text.trim());
                const safeMax = Math.max(1, Math.min(maxResults, 100));

                const response = await context.webAPI.retrieveMultipleRecords(
                    entityName,
                    `?$select=${select}&$filter=${filter}&$top=${safeMax}`,
                );

                const primaryKey = getPrimaryKey(entityName);

                const mapped: SearchResult[] = response.entities.map((entity) => {
                    const id =
                        (entity[primaryKey] as string | undefined) ??
                        (entity["activityid"] as string | undefined) ??
                        "";

                    const displayName =
                        parsedDisplayColumns
                            .map((col) => entity[col] as string | undefined)
                            .filter(Boolean)
                            .join(" · ") || id;

                    const columns = allColumns.reduce<Record<string, string>>((acc, col) => {
                        acc[col] = (entity[col] as string | undefined) ?? "";
                        return acc;
                    }, {});

                    return { id, entityName, columns, displayName };
                });

                setResults(mapped);
                setFocusedIndex(-1);
            } catch (err: unknown) {
                const message =
                    err instanceof Error ? err.message : "Search failed. Please try again.";
                setErrorMessage(message);
                setResults([]);
            } finally {
                setIsLoading(false);
            }
        },
        [context, entityName, maxResults, parsedDisplayColumns, parsedSearchColumns],
    );

    const handleSearchChange = React.useCallback(
        (_ev?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => {
            const val = newValue ?? "";
            setSearchText(val);

            if (debounceTimerRef.current) {
                clearTimeout(debounceTimerRef.current);
            }

            if (val.trim().length === 0) {
                setResults([]);
                setErrorMessage(null);
                return;
            }

            debounceTimerRef.current = setTimeout(() => {
                runSearch(val);
            }, DEBOUNCE_MS);
        },
        [runSearch],
    );

    const handleSelect = React.useCallback(
        (result: SearchResult) => {
            onChange([
                {
                    id: result.id,
                    name: result.displayName,
                    entityType: result.entityName,
                },
            ]);
            setIsSearching(false);
            setSearchText("");
            setResults([]);
            setErrorMessage(null);
        },
        [onChange],
    );

    const handleClear = React.useCallback(() => {
        onChange(undefined);
        setSearchText("");
        setResults([]);
        setErrorMessage(null);
    }, [onChange]);

    const handleDismissSearch = React.useCallback(() => {
        setIsSearching(false);
        setSearchText("");
        setResults([]);
        setErrorMessage(null);
        if (debounceTimerRef.current) {
            clearTimeout(debounceTimerRef.current);
        }
    }, []);

    const handleKeyDown = React.useCallback(
        (ev: React.KeyboardEvent<HTMLElement>) => {
            if (!isCalloutOpen) return;

            if (ev.key === "ArrowDown") {
                ev.preventDefault();
                setFocusedIndex((prev) => Math.min(prev + 1, results.length - 1));
            } else if (ev.key === "ArrowUp") {
                ev.preventDefault();
                setFocusedIndex((prev) => Math.max(prev - 1, -1));
            } else if (ev.key === "Enter") {
                ev.preventDefault();
                if (focusedIndex >= 0 && focusedIndex < results.length) {
                    handleSelect(results[focusedIndex]);
                }
            } else if (ev.key === "Escape") {
                handleDismissSearch();
            }
        },
        [focusedIndex, handleDismissSearch, handleSelect, isCalloutOpen, results],
    );

    // Clean up debounce timer on unmount
    React.useEffect(() => {
        return () => {
            if (debounceTimerRef.current) {
                clearTimeout(debounceTimerRef.current);
            }
        };
    }, []);

    // ── Disabled / read-only display ──────────────────────────────────────────
    if (isDisabled) {
        return (
            <div className={classNames.root}>
                <Text className={classNames.selectedValueText}>
                    {currentDisplayName || "—"}
                </Text>
            </div>
        );
    }

    // ── Selected value display (not in search mode) ───────────────────────────
    if (!isSearching && currentDisplayName) {
        return (
            <div className={classNames.root}>
                <div className={classNames.selectedValueRow}>
                    <Text className={classNames.selectedValueText} title={currentDisplayName}>
                        {currentDisplayName}
                    </Text>
                    <Icon
                        iconName="Edit"
                        className={classNames.iconButton}
                        title="Change selection"
                        onClick={() => setIsSearching(true)}
                        role="button"
                        tabIndex={0}
                        onKeyDown={(e) => e.key === "Enter" && setIsSearching(true)}
                        aria-label="Change selection"
                    />
                    <Icon
                        iconName="Cancel"
                        className={classNames.iconButton}
                        title="Clear selection"
                        onClick={handleClear}
                        role="button"
                        tabIndex={0}
                        onKeyDown={(e) => e.key === "Enter" && handleClear()}
                        aria-label="Clear selection"
                    />
                </div>
            </div>
        );
    }

    // ── Search mode ───────────────────────────────────────────────────────────
    return (
        <div
            className={classNames.root}
            ref={searchBoxContainerRef}
            onKeyDown={handleKeyDown}
        >
            <SearchBox
                placeholder={placeholderText}
                value={searchText}
                onChange={handleSearchChange}
                onClear={handleDismissSearch}
                onFocus={() => setIsSearching(true)}
                autoFocus={isSearching}
                styles={searchBoxStyles}
                aria-label={placeholderText}
                aria-expanded={isCalloutOpen}
                aria-autocomplete="list"
                aria-activedescendant={
                    focusedIndex >= 0 ? `result-item-${focusedIndex}` : undefined
                }
            />

            {isCalloutOpen && searchBoxContainerRef.current && (
                <Callout
                    target={searchBoxContainerRef.current}
                    directionalHint={DirectionalHint.bottomLeftEdge}
                    isBeakVisible={false}
                    onDismiss={handleDismissSearch}
                    calloutWidth={searchBoxContainerRef.current.offsetWidth}
                    preventDismissOnScroll={false}
                    setInitialFocus={false}
                >
                    <div className={classNames.callout} role="listbox">
                        {isLoading && (
                            <Stack
                                horizontal
                                tokens={{ childrenGap: 8 }}
                                verticalAlign="center"
                                style={{ padding: "8px 12px" }}
                            >
                                <Spinner size={SpinnerSize.small} />
                                <Text style={{ fontSize: 13, color: "#605e5c" }}>
                                    Searching…
                                </Text>
                            </Stack>
                        )}

                        {!isLoading && errorMessage && (
                            <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                                {errorMessage}
                            </MessageBar>
                        )}

                        {!isLoading && !errorMessage && results.length === 0 && searchText.trim().length > 0 && (
                            <div className={classNames.noResults}>No results found.</div>
                        )}

                        {!isLoading && !errorMessage && results.length > 0 && (
                            <List
                                items={results}
                                onRenderCell={(item, index) => {
                                    if (!item) return null;
                                    const isFocused = index === focusedIndex;
                                    // Determine secondary display columns (all except the first)
                                    const secondary = parsedDisplayColumns
                                        .slice(1)
                                        .map((col) => item.columns[col])
                                        .filter(Boolean)
                                        .join(" · ");

                                    return (
                                        <div
                                            id={`result-item-${index}`}
                                            role="option"
                                            aria-selected={isFocused}
                                            className={
                                                isFocused
                                                    ? `${classNames.resultItem} ${classNames.resultItemFocused}`
                                                    : classNames.resultItem
                                            }
                                            onClick={() => handleSelect(item)}
                                            onMouseEnter={() => setFocusedIndex(index ?? -1)}
                                            tabIndex={-1}
                                            key={item.id}
                                        >
                                            <Text className={classNames.resultPrimaryText}>
                                                {item.displayName}
                                            </Text>
                                            {secondary && (
                                                <Text className={classNames.resultSecondaryText}>
                                                    {secondary}
                                                </Text>
                                            )}
                                        </div>
                                    );
                                }}
                            />
                        )}
                    </div>
                </Callout>
            )}
        </div>
    );
};
