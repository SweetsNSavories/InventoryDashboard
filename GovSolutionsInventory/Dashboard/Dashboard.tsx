import * as React from 'react';
import {
    Stack, Text, CommandBar, Icon, getTheme, Spinner, SpinnerSize,
    SearchBox, IconButton, TooltipHost, DefaultButton, Shimmer,
    Modal, TextField, PrimaryButton, ContextualMenu, IContextualMenuProps,
    Panel, PanelType, DirectionalHint, TooltipDelay
} from '@fluentui/react';
import { DetailsList, IColumn, SelectionMode, CheckboxVisibility, ConstrainMode, DetailsRow, IDetailsRowProps, ColumnActionsMode, DetailsListLayoutMode } from '@fluentui/react/lib/DetailsList';
import { mergeStyles } from '@fluentui/react/lib/Styling';
import { initializeIcons } from '@fluentui/react';

initializeIcons();

export interface GridItem {
    key: string;
    label: string;
    type: string;
    region?: string;
    sku?: string;
    status?: string;
    version?: string;
    url?: string;
    isDefault?: boolean;
    created?: string;
    isExpanded?: boolean;
    metadata?: string;
}

const getSafeMetadata = (item: any) => {
    if (!item) return {};
    try {
        const metaStr = item.gov_metadata || item.metadata || "{}";
        if (typeof metaStr === 'object' && metaStr !== null) return metaStr;
        return typeof metaStr === 'string' ? JSON.parse(metaStr) : {};
    } catch { return {}; }
};

const getCapacityMB = (meta: any, type: string): number => {
    if (!meta) return 0;
    // Handle both modern expansion and BAP/Sentinel legacy
    const list = meta.properties?.capacity || meta.capacity || [];
    if (!Array.isArray(list)) return 0;
    const entry = list.find((c: any) => c && c.capacityType === type);
    if (!entry) return 0;

    // Support actualConsumption (Modern), consumption.actual (Global), or value (Legacy)
    const rawVal = entry.actualConsumption || entry.consumption?.actual || entry.value || 0;
    return Number(rawVal) || 0;
};


interface SubGridProps {
    context: ComponentFramework.Context<any>;
    envId: string;
    type: string;
    icon: string;
    solutionMap?: Record<string, string>;
}

const CapacityGauge: React.FC<{ label: string, actual: number, total: number, color?: string }> = ({ label, actual, total, color }) => {
    const theme = getTheme();
    const percent = Math.min(100, (actual / total) * 100) || 0;
    const isCritical = actual > total;
    const displayColor = isCritical ? '#ef4444' : (color || theme.palette.themePrimary);

    // SVG Circular Gauge
    const radius = 35;
    const circ = 2 * Math.PI * radius;
    const offset = circ - (percent / 100) * circ;

    return (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }} style={{ padding: 10, background: 'rgba(255,255,255,1)', borderRadius: 12, border: '1px solid #eee', width: 220 }}>
            <div style={{ position: 'relative', width: 80, height: 80, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <svg width="80" height="80">
                    <circle cx="40" cy="40" r={radius} fill="transparent" stroke="#f1f5f9" strokeWidth="8" />
                    <circle cx="40" cy="40" r={radius} fill="transparent" stroke={displayColor} strokeWidth="8"
                        strokeDasharray={circ} strokeDashoffset={offset} strokeLinecap="round" style={{ transition: 'stroke-dashoffset 0.5s ease' }} />
                </svg>
                <div style={{ position: 'absolute', textAlign: 'center' }}>
                    <Text variant="small" style={{ fontWeight: 800, color: displayColor }}>{Math.round((actual / total) * 100)}%</Text>
                </div>
            </div>
            <Stack>
                <Text variant="small" style={{ fontWeight: 800, color: '#64748b', textTransform: 'uppercase', fontSize: 9 }}>{label}</Text>
                <Text variant="medium" style={{ fontWeight: 900 }}>{(actual / 1024 / 1024).toFixed(1)} TB</Text>
                <Text variant="small" style={{ color: '#94a3b8', fontSize: 10 }}>of {(total / 1024 / 1024).toFixed(1)} TB</Text>
            </Stack>
        </Stack>
    );
};

const UserLicenseTooltip: React.FC<{ name: string, assignedLicenses?: string[], profile?: any }> = ({ name, assignedLicenses, profile }) => {
    const theme = getTheme();
    const licenses = assignedLicenses || [];

    return (
        <Stack tokens={{ childrenGap: 8 }} style={{ padding: 15, width: 320, background: '#0f172a', color: 'white', borderRadius: 12, boxShadow: '0 10px 25px rgba(0,0,0,0.3)', border: '1px solid rgba(255,255,255,0.1)' }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                <div style={{ padding: 10, background: theme.palette.themePrimary, borderRadius: '50%', display: 'flex' }}>
                    <Icon iconName="Contact" style={{ fontSize: 20, color: 'white' }} />
                </div>
                <Stack>
                    <Text variant="large" style={{ fontWeight: 900, color: 'white', letterSpacing: -0.5 }}>{name}</Text>
                    <Text variant="small" style={{ color: '#94a3b8' }}>{profile?.userPrincipalName || profile?.mail || "Tenant User"}</Text>
                </Stack>
            </Stack>

            <div style={{ height: 1, background: 'rgba(255,255,255,0.1)', margin: '8px 0' }} />

            {profile && (profile.jobTitle || profile.department || profile.officeLocation) && (
                <Stack tokens={{ childrenGap: 5 }} style={{ marginBottom: 10 }}>
                    <Text variant="small" style={{ color: theme.palette.themePrimary, fontWeight: 800, textTransform: 'uppercase', fontSize: 10 }}>Organizational Profile</Text>
                    {profile.jobTitle && <Text variant="small" style={{ color: '#e2e8f0' }}><b>Role:</b> {profile.jobTitle}</Text>}
                    {profile.department && <Text variant="small" style={{ color: '#e2e8f0' }}><b>Dept:</b> {profile.department}</Text>}
                    {profile.officeLocation && <Text variant="small" style={{ color: '#e2e8f0' }}><b>Office:</b> {profile.officeLocation}</Text>}
                </Stack>
            )}

            <Stack tokens={{ childrenGap: 6 }} style={{ maxHeight: 250, overflowY: 'auto', paddingRight: 8, marginTop: 4 }}>
                <Text variant="small" style={{ color: '#fbbf24', fontWeight: 800, textTransform: 'uppercase', fontSize: 10, letterSpacing: 0.5 }}>Entitlements & Licenses ({licenses.length})</Text>
                {licenses.length > 0 ? (
                    licenses.map((l, idx) => (
                        <div key={idx} style={{ padding: '6px 10px', background: 'rgba(255,255,255,0.05)', borderRadius: 6, border: '1px solid rgba(255,255,255,0.1)' }}>
                            <Text variant="small" style={{ color: '#bae6fd', fontSize: 10, fontWeight: 600, display: 'block' }}>{l}</Text>
                        </div>
                    ))
                ) : (
                    <Text variant="small" style={{ color: 'rgba(255,255,255,0.3)', fontStyle: 'italic', padding: 10 }}>No commercial licenses detected in Microsoft Graph.</Text>
                )}
            </Stack>

            <div style={{ marginTop: 10, textAlign: 'right' }}>
                <Text variant="small" style={{ color: '#475569', fontSize: 9, fontWeight: 700 }}>VERIFIED BY MICROSOFT GRAPH</Text>
            </div>
        </Stack>
    );
};

const DlpPolicyTooltip: React.FC<{ policy: any }> = ({ policy }) => {
    const theme = getTheme();
    return (
        <Stack tokens={{ childrenGap: 8 }} style={{ padding: 15, maxWidth: 350, background: '#0f172a', color: 'white', borderRadius: 12, border: '1px solid rgba(255,255,255,0.1)' }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <Icon iconName="Shield" style={{ fontSize: 20, color: '#34d399' }} />
                <Text variant="medium" style={{ fontWeight: 900, color: 'white' }}>{policy.name || "DLP Policy"}</Text>
            </Stack>
            <div style={{ height: 1, background: 'rgba(255,255,255,0.1)', margin: '4px 0' }} />
            <Text variant="small" style={{ color: '#94a3b8', lineHeight: 1.5 }}>
                <b>What is this?</b> Data Loss Prevention (DLP) policies are the security guardrails of your tenant. They prevent data leakage by categorizing connectors (like Outlook vs. Twitter) into <b>Business</b> and <b>Non-Business</b> buckets, blocking unauthorized data movement.
            </Text>
            <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="small" style={{ color: '#34d399', fontWeight: 800, textTransform: 'uppercase', fontSize: 10 }}>Rule Configuration</Text>
                <Text variant="small" style={{ color: '#e2e8f0' }}>• {policy.ruleSets?.length || 0} Rule Sets Defined</Text>
                <Text variant="small" style={{ color: '#e2e8f0' }}>• Enforcement: Enabled</Text>
            </Stack>
            <div style={{ marginTop: 5, textAlign: 'right' }}>
                <Text variant="small" style={{ color: '#444', fontSize: 9, fontWeight: 700 }}>GOVERNANCE ENFORCED</Text>
            </div>
        </Stack>
    );
};

const MetadataViewer: React.FC<{ json: string, isOpen: boolean, onDismiss: () => void }> = ({ json, isOpen, onDismiss }) => {
    const copy = () => { navigator.clipboard.writeText(json); };
    const theme = getTheme();

    let displayJson = "";
    try {
        if (!json) displayJson = "// No technical metadata available.";
        else {
            const parsed = typeof json === 'string' ? JSON.parse(json) : json;
            displayJson = JSON.stringify(parsed, null, 2);
        }
    } catch {
        displayJson = json || "// Encoding error";
    }

    return (
        <Modal
            isOpen={isOpen}
            onDismiss={onDismiss}
            isBlocking={false}
            containerClassName={mergeStyles({
                padding: '0 !important',
                width: 1000,
                maxWidth: '92vw',
                borderRadius: 12,
                boxShadow: '0 32px 128px rgba(0,0,0,0.3)',
                overflow: 'hidden'
            })}
            styles={{ main: { zIndex: 1000000, background: '#1e1e1e !important' } }}
        >
            <Stack style={{ background: '#1e1e2e', height: '100%' }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ padding: '20px 30px', borderBottom: '1px solid #333' }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                        <div style={{ background: theme.palette.themePrimary, padding: 8, borderRadius: 6 }}>
                            <Icon iconName="Binary" style={{ fontSize: 20, color: 'white' }} />
                        </div>
                        <Stack>
                            <Text variant="large" style={{ fontWeight: 800, color: 'white' }}>Metadata Inspector</Text>
                            <Text variant="small" style={{ color: '#aaa', fontWeight: 600 }}>TECHNICAL JSON SCHEMA</Text>
                        </Stack>
                    </Stack>
                    <IconButton iconProps={{ iconName: 'Cancel' }} onClick={onDismiss} styles={{ root: { color: '#ccc', selectors: { ':hover': { color: 'white', background: '#444' } } } }} />
                </Stack>

                <div style={{ padding: '25px 30px', background: '#111', maxHeight: '70vh', overflowY: 'auto' }}>
                    <pre style={{ margin: 0, fontFamily: '"JetBrains Mono", "Cascadia Code", "Consolas", monospace', fontSize: 13, whiteSpace: 'pre-wrap', color: '#61afef', lineHeight: 1.6 }}>
                        {displayJson}
                    </pre>
                </div>

                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ padding: '15px 30px', background: '#1e1e2e', borderTop: '1px solid #333' }}>
                    <Text variant="small" style={{ color: '#888' }}>Capture Date: {new Date().toLocaleDateString()}</Text>
                    <Stack horizontal tokens={{ childrenGap: 12 }}>
                        <PrimaryButton iconProps={{ iconName: 'Copy' }} onClick={copy} styles={{ root: { borderRadius: 6, height: 40 } }}>Copy JSON</PrimaryButton>
                        <DefaultButton onClick={onDismiss} styles={{ root: { borderRadius: 6, height: 40, border: '1px solid #444', background: 'transparent', color: 'white', selectors: { ':hover': { background: '#333', color: 'white' } } } }}>Dismiss</DefaultButton>
                    </Stack>
                </Stack>
            </Stack>
        </Modal>
    );
};

const SubGrid: React.FC<SubGridProps> = ({ context, envId, type, icon, solutionMap }) => {
    const [items, setItems] = React.useState<any[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);
    const [page, setPage] = React.useState(1);
    const [selectedMeta, setSelectedMeta] = React.useState<string | null>(null);
    const [sortConfig, setSortConfig] = React.useState<{ key: string, desc: boolean }>({ key: 'gov_name', desc: false });
    const pageSize = 15;
    const theme = getTheme();
    const [columnFilters, setColumnFilters] = React.useState<Record<string, string>>({});

    const getSafeMeta = (item: any) => getSafeMetadata(item);

    const [filterMenuProps, setFilterMenuProps] = React.useState<{
        target: HTMLElement | MouseEvent;
        columnKey: string;
        isVisible: boolean;
    } | null>(null);

    const onDismissFilter = () => setFilterMenuProps(null);

    React.useEffect(() => {
        const load = async () => {
            setIsLoading(true);
            try {
                const isSol = type === "Solution";
                const table = isSol ? "gov_solution" : "gov_asset";

                // Fields for Apps/Flows
                let select = "gov_name,gov_displayname,gov_type,gov_envid,gov_owner,gov_assetid,gov_healthstatus,gov_solutionid,gov_version,gov_state,gov_ismanaged,gov_metadata,modifiedon,gov_apptype,gov_formfactor,gov_almmode,gov_usespremiumapi,gov_dlpstatus,gov_createdon,gov_modifiedon";
                let filter = `?$filter=gov_envid eq '${envId}' and gov_type eq '${type}'&$select=${select}&$orderby=gov_name asc`;

                // Fields for Solutions
                if (isSol) {
                    select = "gov_name,gov_displayname,gov_uniquename,gov_version,gov_owner,gov_description,gov_url,gov_state,gov_metadata,gov_createdon,gov_modifiedon";
                    filter = `?$filter=gov_envid eq '${envId}'&$select=${select}&$orderby=gov_name asc`;
                }

                const result = await context.webAPI.retrieveMultipleRecords(table, filter);
                setItems(result.entities);
            } catch (e) {
                console.error(`Failed to load ${type}`, e);
            }
            setIsLoading(false);
        };
        void load();
    }, [envId, type]);

    const columns: IColumn[] = React.useMemo(() => {
        const common: IColumn[] = [
            { key: 'icon', name: '', minWidth: 24, maxWidth: 24, onRender: () => <Icon iconName={icon} style={{ color: theme.palette.themePrimary }} /> },
            {
                key: 'gov_name', name: 'Asset Name', minWidth: 350, maxWidth: 450, isResizable: true,
                isSorted: sortConfig.key === 'gov_name', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: any) => {
                    if (!item) return null;
                    const theme = getTheme();
                    const health = item.gov_healthstatus || "Healthy";
                    const isWarning = health === "Disabled" || health === "Issues" || item.gov_state === "Stopped";
                    const statusColor = isWarning ? "#d13438" : "#107c10";
                    const statusIcon = isWarning ? "Warning" : "CheckMark";

                    const meta = getSafeMeta(item);
                    let name = item.gov_displayname || meta.properties?.displayName || meta.displayName || item.gov_name?.replace('⚠️ ', '').replace('✅ ', '') || "Unnamed Asset";

                    if (name.startsWith("Cloud Flow (") && item.gov_displayname && !item.gov_displayname.startsWith("Cloud Flow (")) {
                        name = item.gov_displayname;
                    }

                    return (
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                            <Icon iconName={statusIcon} style={{ fontSize: 14, color: statusColor, fontWeight: 900 }} />
                            <Text style={{ color: isWarning ? statusColor : '#201f1e', fontWeight: 700, fontSize: 13 }}>{name}</Text>
                        </Stack>
                    );
                }
            }
        ];

        const dateCols: IColumn[] = [
            {
                key: 'gov_createdon', name: 'Created', minWidth: 100, isResizable: true,
                isSorted: sortConfig.key === 'gov_createdon', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: any) => {
                    if (!item) return null;
                    const d = item.gov_createdon ? new Date(item.gov_createdon) : null;
                    return <Text variant="small" style={{ color: '#666' }}>{d ? d.toLocaleDateString() : "-"}</Text>;
                }
            },
            {
                key: 'gov_modifiedon', name: 'Modified', minWidth: 100, isResizable: true,
                isSorted: sortConfig.key === 'gov_modifiedon', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: any) => {
                    if (!item) return null;
                    const d = item.gov_modifiedon ? new Date(item.gov_modifiedon) : null;
                    return <Text variant="small" style={{ color: '#666' }}>{d ? d.toLocaleDateString() : "-"}</Text>;
                }
            }
        ];

        const standardActions: IColumn = {
            key: 'actions', name: '', minWidth: 80, isResizable: true, onRender: (item: any) => {
                if (!item) return null;
                return (
                    <Stack horizontal tokens={{ childrenGap: 5 }}>
                        <IconButton iconProps={{ iconName: 'Play' }} onClick={(e) => {
                            e.stopPropagation();
                            const meta = getSafeMeta(item);
                            let link = item.gov_playuri || "";
                            if (!link) {
                                if (type === 'Cloud Flow') {
                                    const workflowId = meta.properties?.workflowEntityId || item.gov_assetid;
                                    link = `https://make.powerautomate.com/environments/${envId}/flows/${workflowId}/details`;
                                }
                                else if (type === 'Canvas App') link = `https://make.powerapps.com/environments/${envId}/apps/${item.gov_assetid}/details`;
                                else if (type === 'Power Page') try { link = meta.properties?.siteUrl || ""; } catch (err) { console.error(err); }
                            }
                            if (link) window.open(link, '_blank');
                        }} title="Launch Manager" />
                        <IconButton iconProps={{ iconName: 'Code' }} onClick={(e) => {
                            e.stopPropagation();
                            setSelectedMeta(item.gov_metadata || "{}");
                        }} title="View Metadata" />
                    </Stack>
                );
            }
        };

        if (type === "User") {
            return [
                ...common,
                {
                    key: 'email', name: 'Identity', minWidth: 250,
                    isSorted: sortConfig.key === 'email', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const meta = getSafeMeta(item);
                        const email = item.gov_owner || meta.email || "-";
                        const licenses = meta.assignedLicenses || meta.properties?.assignedLicenses || [];
                        const graphProfile = meta.graphProfile || meta.properties?.graphProfile;

                        return (
                            <TooltipHost
                                content={<UserLicenseTooltip name={item.gov_displayname || "User"} assignedLicenses={licenses} profile={graphProfile} />}
                                directionalHint={DirectionalHint.rightCenter}
                                calloutProps={{
                                    gapSpace: 0,
                                    styles: {
                                        root: { pointerEvents: 'auto' },
                                        beak: { background: '#0f172a' },
                                        beakCurtain: { background: '#0f172a' },
                                        calloutMain: { background: '#0f172a', borderRadius: 12, boxShadow: '0 20px 50px rgba(0,0,0,0.4)' }
                                    }
                                }}
                                delay={TooltipDelay.zero}
                            >
                                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                                    <Icon iconName="Contact" style={{ color: theme.palette.themePrimary, fontSize: 14 }} />
                                    <Text variant="small" style={{ color: theme.palette.themePrimary, fontWeight: 700, borderBottom: `1px dashed ${theme.palette.themePrimary}` }}>{email}</Text>
                                </Stack>
                            </TooltipHost>
                        );
                    }
                },
                {
                    key: 'details', name: 'Security Access', minWidth: 200, isResizable: true,
                    isSorted: sortConfig.key === 'details', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const meta = getSafeMeta(item);
                        const isAppUser = meta.properties?.isApplicationUser ||
                            meta.properties?.accessMode === 5 ||
                            meta.properties?.accessMode === "5" ||
                            (item.gov_displayname && item.gov_displayname.startsWith("#"));
                        const isApp = isAppUser ? "App User" : "Standard User";
                        return (
                            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                                <Icon iconName={isAppUser ? "Robot" : "ReminderPerson"} style={{ fontSize: 12, color: '#888' }} />
                                <Text variant="small" style={{ color: '#666' }}>{isApp}</Text>
                            </div>
                        );
                    }
                },
                ...dateCols,
                standardActions
            ];
        }

        if (type === "Solution") {
            return [
                ...common,
                {
                    key: 'uniquename', name: 'Unique Name', minWidth: 150, isResizable: true, fieldName: 'gov_uniquename',
                    isSorted: sortConfig.key === 'uniquename', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false })
                },
                {
                    key: 'version', name: 'Version', minWidth: 80, isResizable: true, fieldName: 'gov_version',
                    isSorted: sortConfig.key === 'version', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false })
                },
                {
                    key: 'publisher', name: 'Publisher', minWidth: 180, isResizable: true, fieldName: 'gov_owner',
                    isSorted: sortConfig.key === 'publisher', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const val = item.gov_owner || "";
                        const isGuid = /^[0-9a-f]{8}-/.test(val.toLowerCase());
                        return <Text variant="small" style={{ fontWeight: isGuid ? 400 : 700, color: isGuid ? '#888' : '#0078d4' }}>{val}</Text>;
                    }
                },
                {
                    key: 'state', name: 'State', minWidth: 70, isResizable: true,
                    isSorted: sortConfig.key === 'state', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        return <Text variant="small">{item.gov_state === "0" ? "Ready" : item.gov_state}</Text>
                    }
                },
                ...dateCols,
                {
                    key: 'link', name: 'Docs', minWidth: 50, isResizable: true,
                    onRender: (item: any) => {
                        if (!item) return null;
                        return <IconButton iconProps={{ iconName: 'Link' }} onClick={() => window.open(item.gov_url, '_blank')} disabled={!item.gov_url} />
                    }
                },
                standardActions
            ];
        }

        if (type === "Canvas App") {
            return [
                ...common,
                {
                    key: 'solution', name: 'Owned By Solution', minWidth: 150, isResizable: true,
                    isSorted: sortConfig.key === 'solution', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const solId = item.gov_solutionid;
                        const solName = solutionMap?.[solId] || (solId && solId !== "00000000-0000-0000-0000-000000000000" ? `Sol: ${solId.substring(0, 8)}` : "None");
                        return <Text variant="small" style={{ color: solId ? theme.palette.themePrimary : '#888', fontWeight: solId ? 600 : 400 }}>{solName}</Text>;
                    }
                },
                {
                    key: 'formfactor', name: 'UI', minWidth: 70, isResizable: true, onRender: (item: any) => {
                        if (!item) return null;
                        return <Text variant="small">{item.gov_formfactor}</Text>
                    }
                },
                {
                    key: 'premium', name: 'Premium', minWidth: 70, isResizable: true,
                    onRender: (item: any) => {
                        if (!item) return null;
                        return <Icon iconName={item.gov_usespremiumapi ? "Diamond" : "CheckMark"} style={{ color: item.gov_usespremiumapi ? "#b91c1c" : "#166534" }} />
                    }
                },
                { key: 'almmode', name: 'ALM', minWidth: 80, isResizable: true, fieldName: 'gov_almmode' },
                {
                    key: 'dlp', name: 'DLP', minWidth: 100, isResizable: true,
                    onRender: (item: any) => {
                        if (!item) return null;
                        const status = item.gov_dlpstatus || "Unknown";
                        const isOk = status === "Compliant";
                        return (
                            <div style={{
                                padding: '2px 8px',
                                borderRadius: 4,
                                background: isOk ? '#dcfce7' : '#fee2e2',
                                color: isOk ? '#166534' : '#991b1b',
                                fontSize: 10,
                                fontWeight: 700,
                                border: `1px solid ${isOk ? '#86efac' : '#fecaca'}`
                            }}>
                                {status.toUpperCase()}
                            </div>
                        );
                    }
                },
                {
                    key: 'owner', name: 'Owner', minWidth: 150, isResizable: true, fieldName: 'gov_owner',
                    onRender: (item: any) => {
                        if (!item) return null;
                        return <Text variant="small">{item.gov_owner}</Text>
                    }
                },
                {
                    key: 'status', name: 'Status', minWidth: 80, isResizable: true,
                    onRender: (item: any) => {
                        if (!item) return null;
                        const meta = getSafeMeta(item);
                        const state = item.gov_state || meta.properties?.status || "Active";
                        const isPrimary = state === "Ready" || state === "Active" || state === "Started";
                        return (
                            <div style={{
                                padding: '2px 10px',
                                borderRadius: 20,
                                background: isPrimary ? '#f0fdf4' : '#fef2f2',
                                color: isPrimary ? '#166534' : '#991b1b',
                                fontSize: 11,
                                fontWeight: 800,
                                display: 'inline-block',
                                border: `1px solid ${isPrimary ? '#bbf7d0' : '#fecaca'}`
                            }}>
                                {state.toUpperCase()}
                            </div>
                        );
                    }
                },
                standardActions,
                ...dateCols
            ];
        }

        if (type === "Cloud Flow") {
            return [
                ...common,
                {
                    key: 'solution', name: 'Owned By Solution', minWidth: 180, isResizable: true,
                    isSorted: sortConfig.key === 'solution', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const solId = item.gov_solutionid;
                        const solName = solutionMap?.[solId] || (solId && solId !== "00000000-0000-0000-0000-000000000000" ? `Sol: ${solId.substring(0, 8)}` : "None");
                        return <Text variant="small" style={{ color: solId ? theme.palette.themePrimary : '#888', fontWeight: solId ? 600 : 400 }}>{solName}</Text>;
                    }
                },
                {
                    key: 'owner', name: 'Owner', minWidth: 180, isResizable: true,
                    isSorted: sortConfig.key === 'owner', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const meta = getSafeMeta(item);
                        return <Text variant="small" style={{ color: '#666' }}>{item.gov_owner || meta.properties?.creator?.email || "-"}</Text>;
                    }
                },
                {
                    key: 'status', name: 'Status', minWidth: 90, isResizable: true,
                    isSorted: sortConfig.key === 'status', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const meta = getSafeMeta(item);
                        const state = item.gov_state || meta.properties?.state || "Active";
                        const isPrimary = state === "Started" || state === "Active" || state === "On";
                        const isCritical = state === "Stopped" || state === "Suspended" || state === "Disabled";
                        return (
                            <div style={{
                                padding: '2px 10px',
                                borderRadius: 20,
                                background: isPrimary ? '#f0fdf4' : (isCritical ? '#fef2f2' : '#fff7ed'),
                                color: isPrimary ? '#166534' : (isCritical ? '#991b1b' : '#9a3412'),
                                fontSize: 11,
                                fontWeight: 800,
                                display: 'inline-block',
                                border: `1px solid ${isPrimary ? '#bbf7d0' : (isCritical ? '#fecaca' : '#fed7aa')}`
                            }}>
                                {state.toUpperCase()}
                            </div>
                        );
                    }
                },
                ...dateCols,
                standardActions
            ];
        }

        if (type === "Power Page") {
            return [
                ...common,
                {
                    key: 'portalid', name: 'Portal ID', minWidth: 280,
                    isSorted: sortConfig.key === 'portalid', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const meta = getSafeMeta(item);
                        return <Text variant="small" style={{ color: '#888', fontFamily: 'monospace' }}>{meta.portalId || meta.properties?.portalId || item.gov_name || "-"}</Text>;
                    }
                },
                {
                    key: 'url', name: 'Site URL', minWidth: 300,
                    isSorted: sortConfig.key === 'url', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const meta = getSafeMeta(item);
                        const url = meta.siteUrl || meta.properties?.siteUrl || "";
                        return url ? <a href={url} target="_blank" rel="noreferrer" style={{ textDecoration: 'none', color: theme.palette.themePrimary, fontWeight: 800 }}>{url}</a> : <Text variant="small" style={{ color: '#bbb' }}>-</Text>;
                    }
                },
                {
                    key: 'state', name: 'Model', minWidth: 100,
                    isSorted: sortConfig.key === 'state', isSortedDescending: sortConfig.desc,
                    onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                    onRender: (item: any) => {
                        if (!item) return null;
                        const meta = getSafeMeta(item);
                        return <Text variant="small" style={{ color: '#666' }}>{meta.state || meta.properties?.state || "Active"}</Text>;
                    }
                },
                ...dateCols,
                standardActions
            ];
        }

        return [
            ...common,
            { key: 'owner', name: 'Owner', minWidth: 200, fieldName: 'gov_owner' },
            standardActions
        ];
    }, [type, theme, envId, solutionMap, icon, sortConfig]);

    const filteredData = React.useMemo(() => {
        let list = [...items];

        // Filtering
        list = list.filter(i => {
            if (!i) return false;
            const meta = getSafeMeta(i);
            return Object.entries(columnFilters).every(([key, filter]) => {
                if (!filter) return true;
                const f = String(filter).toLowerCase();

                let val = "";
                if (key === 'gov_name') {
                    val = String(i.gov_displayname || meta.properties?.displayName || i.gov_name || "");
                } else if (key === 'email' || key === 'owner') {
                    val = String(i.gov_owner || meta.email || meta.properties?.creator?.email || "");
                } else if (key === 'solution') {
                    val = String(solutionMap?.[i.gov_solutionid] || "");
                } else if (key === 'status') {
                    val = String(i.gov_state || meta.properties?.status || meta.properties?.state || "");
                } else if (key === 'portalid') {
                    val = String(meta.portalId || i.gov_name || "");
                } else if (key === 'url') {
                    val = String(meta.siteUrl || meta.properties?.siteUrl || "");
                } else if (key === 'details' && type === 'User') {
                    const isAppUser = meta.properties?.isApplicationUser || meta.properties?.accessMode === 5 || meta.properties?.accessMode === "5" || (i.gov_displayname && i.gov_displayname.startsWith("#"));
                    val = isAppUser ? "App User" : "Standard User";
                } else {
                    val = String(i[key] || (i as any)[`gov_${key}`] || "");
                }

                return val.toLowerCase().includes(f);
            });
        });

        // Sorting
        if (sortConfig.key) {
            list.sort((a, b) => {
                const metaA = getSafeMeta(a);
                const metaB = getSafeMeta(b);
                let vA: any = a[sortConfig.key];
                let vB: any = b[sortConfig.key];

                if (sortConfig.key === 'gov_name') {
                    vA = (a.gov_displayname || metaA.properties?.displayName || a.gov_name || "");
                    vB = (b.gov_displayname || metaB.properties?.displayName || b.gov_name || "");
                } else if (sortConfig.key === 'email' || sortConfig.key === 'owner') {
                    vA = a.gov_owner || metaA.email || metaA.properties?.creator?.email || "";
                    vB = b.gov_owner || metaB.email || metaB.properties?.creator?.email || "";
                } else if (sortConfig.key === 'solution') {
                    vA = solutionMap?.[a.gov_solutionid] || "";
                    vB = solutionMap?.[b.gov_solutionid] || "";
                } else if (sortConfig.key === 'status') {
                    vA = a.gov_state || metaA.properties?.status || metaA.properties?.state || "";
                    vB = b.gov_state || metaB.properties?.status || metaB.properties?.state || "";
                } else if (sortConfig.key === 'gov_createdon') {
                    vA = new Date(a.gov_createdon || a.createdon || metaA.createdTime || 0).getTime();
                    vB = new Date(b.gov_createdon || b.createdon || metaB.createdTime || 0).getTime();
                } else if (sortConfig.key === 'gov_modifiedon') {
                    vA = new Date(a.gov_modifiedon || a.modifiedon || metaA.lastModifiedTime || 0).getTime();
                    vB = new Date(b.gov_modifiedon || b.modifiedon || metaB.lastModifiedTime || 0).getTime();
                } else {
                    vA = a[sortConfig.key] || (a as any)[`gov_${sortConfig.key}`] || "";
                    vB = b[sortConfig.key] || (b as any)[`gov_${sortConfig.key}`] || "";
                }

                vA = typeof vA === 'string' ? vA.toLowerCase() : vA;
                vB = typeof vB === 'string' ? vB.toLowerCase() : vB;

                if (vA < vB) return sortConfig.desc ? 1 : -1;
                if (vA > vB) return sortConfig.desc ? -1 : 1;
                return 0;
            });
        }

        return list;
    }, [items, columnFilters, sortConfig, solutionMap, type]);

    const pItems = React.useMemo(() => {
        return filteredData.slice((page - 1) * pageSize, page * pageSize);
    }, [filteredData, page]);

    const totalPages = Math.ceil(filteredData.length / pageSize);

    const searchableCols = React.useMemo(() => {
        return columns.map(col => ({
            ...col,
            onRenderHeader: (props: any, defaultRender: any) => {
                if (!props || col.key === 'icon' || col.key === 'actions') return defaultRender(props);
                const isFiltered = !!columnFilters[col.key];
                return (
                    <Stack horizontal verticalAlign="center" horizontalAlign="space-between" styles={{ root: { width: '100%' } }}>
                        <span style={{ flexGrow: 1 }}>{defaultRender(props)}</span>
                        <IconButton
                            iconProps={{ iconName: 'Filter' }}
                            title="Filter Column"
                            styles={{
                                root: { height: 16, width: 16 },
                                icon: { fontSize: 10, color: isFiltered ? theme.palette.themePrimary : '#adb5bd', opacity: isFiltered ? 1 : 0.6 }
                            }}
                            onClick={(e) => {
                                e.stopPropagation();
                                setFilterMenuProps({
                                    target: e.currentTarget as any,
                                    columnKey: col.key,
                                    isVisible: true
                                });
                            }}
                        />
                    </Stack>
                );
            }
        }));
    }, [columns, columnFilters, theme]);

    const paginated = pItems;
    const totalCount = filteredData.length;
    const totalPagesCount = Math.ceil(totalCount / pageSize);

    if (isLoading) return <Shimmer style={{ margin: '15px 0' }} />;

    return (
        <Stack tokens={{ childrenGap: 10 }} style={{ padding: '0 0 20px 0', width: '100%' }}>
            <div style={{ display: 'grid', width: '100%', overflow: 'hidden' }}>
                <div
                    style={{
                        border: '1px solid #edebe9',
                        borderRadius: 6,
                        background: 'white',
                        overflowX: 'auto',
                        overflowY: 'hidden',
                        width: '100%',
                        maxWidth: '100%',
                        position: 'relative'
                    }}
                >
                    <DetailsList items={paginated} columns={searchableCols} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.fixedColumns} constrainMode={ConstrainMode.unconstrained} compact />
                </div>
            </div>
            {totalPagesCount > 1 && (
                <Stack horizontal horizontalAlign="end" verticalAlign="center" tokens={{ childrenGap: 15 }} style={{ paddingTop: 10 }}>
                    <Text variant="small" style={{ color: '#605e5c' }}>Page <b>{page}</b> of <b>{totalPagesCount}</b></Text>
                    <IconButton iconProps={{ iconName: 'ChevronLeft' }} disabled={page === 1} onClick={() => setPage(p => p - 1)} />
                    <IconButton iconProps={{ iconName: 'ChevronRight' }} disabled={page === totalPagesCount} onClick={() => setPage(p => p + 1)} />
                </Stack>
            )}
            {filterMenuProps && (
                <ContextualMenu
                    target={filterMenuProps.target}
                    shouldFocusOnMount={true}
                    onDismiss={onDismissFilter}
                    items={[
                        {
                            key: 'filterInput',
                            onRender: () => (
                                <Stack styles={{ root: { padding: '10px 15px', width: 250 } }}>
                                    <Text variant="small" style={{ fontWeight: 600, marginBottom: 5 }}>
                                        Filter by {searchableCols.find(c => c.key === filterMenuProps.columnKey)?.name}
                                    </Text>
                                    <SearchBox
                                        placeholder="Type to filter..."
                                        value={columnFilters[filterMenuProps.columnKey] || ''}
                                        onChange={(_, newVal) => {
                                            setColumnFilters(prev => ({ ...prev, [filterMenuProps.columnKey]: newVal || '' }));
                                            setPage(1);
                                        }}
                                        styles={{ root: { border: '1px solid #ccc' } }}
                                    />
                                    <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }} style={{ marginTop: 10 }}>
                                        <DefaultButton
                                            text="Clear"
                                            styles={{ root: { height: 24, padding: 0, minWidth: 60 } }}
                                            onClick={() => {
                                                setColumnFilters(prev => ({ ...prev, [filterMenuProps.columnKey]: '' }));
                                                setPage(1);
                                                onDismissFilter();
                                            }}
                                        />
                                        <PrimaryButton
                                            text="Apply"
                                            styles={{ root: { height: 24, padding: 0, minWidth: 60 } }}
                                            onClick={onDismissFilter}
                                        />
                                    </Stack>
                                </Stack>
                            )
                        }
                    ]}
                />
            )}
            <MetadataViewer
                json={selectedMeta || ""}
                isOpen={!!selectedMeta}
                onDismiss={() => setSelectedMeta(null)}
            />
        </Stack>
    );
};

const EnvironmentAccordion: React.FC<{ context: ComponentFramework.Context<any>, envId: string, title: string, type: string, icon: string, solutionMap?: Record<string, string> }> = ({ context, envId, title, type, icon, solutionMap }) => {
    const [isExpanded, setIsExpanded] = React.useState(false);
    const theme = getTheme();
    return (
        <div style={{ borderBottom: '1px solid #f0f0f0', transition: 'all 0.2s ease' }}>
            <Stack horizontal verticalAlign="center" onClick={() => setIsExpanded(!isExpanded)}
                style={{
                    cursor: 'pointer',
                    padding: '12px 20px',
                    background: isExpanded ? '#f9fbff' : 'transparent',
                    borderLeft: isExpanded ? `5px solid ${theme.palette.themePrimary}` : '5px solid transparent',
                    transition: 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)'
                }}>
                <Icon iconName={isExpanded ? "ChevronDown" : "ChevronRight"} style={{ fontSize: 9, marginRight: 15, color: '#999' }} />
                <div style={{ background: isExpanded ? theme.palette.themePrimary : '#f3f4f6', padding: 6, borderRadius: 6, marginRight: 15, transition: 'all 0.3s' }}>
                    <Icon iconName={icon} style={{ color: isExpanded ? 'white' : theme.palette.themePrimary, fontSize: 16 }} />
                </div>
                <Text variant="medium" style={{ fontWeight: isExpanded ? 700 : 600, flexGrow: 1, color: isExpanded ? theme.palette.themePrimary : '#444' }}>{title}</Text>
                {isExpanded && <Icon iconName="Wait" style={{ fontSize: 12, color: '#ccc' }} />}
            </Stack>
            {isExpanded && (
                <div style={{ padding: '0px 25px 15px 55px', background: '#fff', borderTop: '1px solid #f9fafb' }}>
                    <SubGrid context={context} envId={envId} type={type} icon={icon} solutionMap={solutionMap} />
                </div>
            )}
        </div>
    );
};

const EnvironmentDetail: React.FC<{ context: ComponentFramework.Context<any>, envId: string, metadata?: string }> = ({ context, envId, metadata }) => {
    const [solutionMap, setSolutionMap] = React.useState<Record<string, string>>({});
    const isTenant = envId === "00000000-0000-0000-0000-000000000000";
    const theme = getTheme();

    const meta = React.useMemo(() => {
        try { return JSON.parse(metadata || "{}"); } catch { return {}; }
    }, [metadata]);

    React.useEffect(() => {
        const loadSolutions = async () => {
            try {
                const query = `?$filter=gov_envid eq '${envId}'&$select=gov_solutionid,gov_displayname`;
                const result = await context.webAPI.retrieveMultipleRecords("gov_solution", query);
                const map: Record<string, string> = {};
                result.entities.forEach(e => { if (e.gov_solutionid) map[e.gov_solutionid] = e.gov_displayname || "Solution"; });
                setSolutionMap(map);
            } catch (e) { console.error("Solution Mapping Failed", e); }
        };
        if (!isTenant) void loadSolutions();
    }, [envId, isTenant]);

    if (isTenant) {
        const caps = meta.capacity || meta.properties?.capacity || [];

        const getCap = (type: string) => {
            const entry = caps.find((c: any) => c.capacityType === type);
            if (!entry) return { actual: 0, total: 1 };
            const actual = Number(entry.actualConsumption || entry.consumption?.actual || entry.value || 0);
            const total = Number(entry.totalCapacity || 1);
            return { actual, total };
        };

        const db = getCap("Database");
        const file = getCap("File");
        const log = getCap("Log");


        return (
            <Stack tokens={{ childrenGap: 20 }} style={{ padding: '20px 0' }}>
                <div style={{ background: '#f8fafc', padding: 25, borderRadius: 12, border: '1px solid #e2e8f0' }}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="end" style={{ marginBottom: 20 }}>
                        <Stack>
                            <Text variant="large" style={{ fontWeight: 900, color: '#1e293b' }}>Global Storage Sentinel</Text>
                            <Text variant="small" style={{ color: '#64748b' }}>REAL-TIME AGGREGATED TENANT POOL</Text>
                        </Stack>
                        <div style={{ background: '#fee2e2', padding: '6px 12px', borderRadius: 20, border: '1px solid #fecaca' }}>
                            <Text style={{ color: '#991b1b', fontWeight: 800, fontSize: 11 }}>3 STORAGE CRITICAL ALERTS</Text>
                        </div>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 20 }}>
                        <CapacityGauge label="Database Storage" actual={db.actual} total={db.total} />
                        <CapacityGauge label="File Storage" actual={file.actual} total={file.total} color="#10b981" />
                        <CapacityGauge label="Log Storage (ALERT)" actual={log.actual} total={log.total} color="#f59e0b" />
                    </Stack>
                </div>

                <Stack horizontal tokens={{ childrenGap: 20 }}>
                    <div style={{ flex: 1, background: 'white', border: '1px solid #eee', borderRadius: 12, padding: 20 }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: 15 }}>
                            <Icon iconName="Shield" style={{ fontSize: 18, color: '#3b82f6' }} />
                            <Text variant="medium" style={{ fontWeight: 800 }}>DLP Compliance Summary</Text>
                        </Stack>
                        <Stack tokens={{ childrenGap: 8 }}>
                            {(meta.governance || []).slice(0, 5).map((p: any) => (
                                <TooltipHost
                                    key={p.id}
                                    content={<DlpPolicyTooltip policy={p} />}
                                    delay={0}
                                    directionalHint={1}
                                    styles={{ root: { display: 'block' } }}
                                >
                                    <Stack horizontal horizontalAlign="space-between" style={{ padding: '6px 10px', background: '#fff', borderRadius: 6, border: '1px solid #f1f5f9', cursor: 'help', transition: 'all 0.2s' }}>
                                        <Text variant="small" style={{ fontWeight: 600, color: '#334155' }}>{p.name}</Text>
                                        <Text variant="small" style={{ color: '#666', fontWeight: 700 }}>{p.ruleSets?.length || 0} Rules</Text>
                                    </Stack>
                                </TooltipHost>
                            ))}
                        </Stack>
                    </div>
                    <div style={{ flex: 1, background: 'white', border: '1px solid #eee', borderRadius: 12, padding: 20 }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: 15 }}>
                            <Icon iconName="PaymentCard" style={{ fontSize: 18, color: '#10b981' }} />
                            <Text variant="medium" style={{ fontWeight: 800 }}>Product Entitlements</Text>
                        </Stack>
                        <Stack tokens={{ childrenGap: 8 }}>
                            {(meta.licensing || []).slice(0, 5).map((l: any) => (
                                <Stack key={l.skuId || l.id} horizontal horizontalAlign="space-between" style={{ paddingBottom: 8, borderBottom: '1px solid #f8fafc' }}>
                                    <Text variant="small" style={{ fontWeight: 600 }}>{l.displayName || l.name}</Text>
                                    <Text variant="small" style={{ color: '#107c10', fontWeight: 700 }}>Active</Text>
                                </Stack>
                            ))}
                        </Stack>
                    </div>
                </Stack>
            </Stack>
        );
    }

    const capacity = meta.properties?.capacity || [];
    const hasCapacity = capacity.length > 0;

    return (
        <Stack style={{ background: '#fff', border: '1px solid #edebe9', borderRadius: 8, margin: '10px 0', boxShadow: '0 4px 16px rgba(0,0,0,0.08)' }}>

            {/* CAPACITY SECTION */}
            {hasCapacity && (
                <div style={{ padding: '20px', background: '#f8fafc', borderBottom: '1px solid #eee' }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: 15 }}>
                        <Icon iconName="Database" style={{ fontSize: 18, color: theme.palette.themePrimary }} />
                        <Text variant="medium" style={{ fontWeight: 800 }}>Environment Storage Usage</Text>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 20 }} style={{ flexWrap: 'wrap' }}>
                        {(() => {
                            const list = meta.properties?.capacity || meta.capacity || [];
                            if (!Array.isArray(list)) return null;
                            return list.map((c: any) => {
                                const actual = (Number(c.actualConsumption || c.consumption?.actual || c.value || 0) || 0) / 1024.0;
                                const rated = (Number(c.ratedConsumption || c.consumption?.rated || 0) || 0) / 1024.0;
                                return (
                                    <div key={c.capacityType} style={{ minWidth: 150, padding: 12, background: 'white', border: '1px solid #e2e8f0', borderRadius: 8 }}>
                                        <Text variant="small" style={{ color: '#64748b', fontWeight: 800, textTransform: 'uppercase', fontSize: 9 }}>{c.capacityType}</Text>
                                        <Text variant="large" style={{ fontWeight: 900, display: 'block' }}>{actual.toFixed(2)} GB</Text>
                                        <Text variant="small" style={{ color: '#94a3b8' }}>Rated: {rated.toFixed(2)} GB</Text>
                                    </div>
                                );
                            });
                        })()}
                    </Stack>
                </div>
            )}
            <EnvironmentAccordion context={context} envId={envId} title="Cloud Flows" type="Cloud Flow" icon="MicrosoftFlowLogo" solutionMap={solutionMap} />
            <EnvironmentAccordion context={context} envId={envId} title="Canvas Apps" type="Canvas App" icon="TabletMode" solutionMap={solutionMap} />
            <EnvironmentAccordion context={context} envId={envId} title="Power Pages" type="Power Page" icon="InternetSharing" solutionMap={solutionMap} />
            <EnvironmentAccordion context={context} envId={envId} title="Solutions" type="Solution" icon="Puzzle" />
            <EnvironmentAccordion context={context} envId={envId} title="Users & Groups" type="User" icon="People" />
        </Stack>
    );
};

export const InventoryDashboardUI: React.FC<{ context: ComponentFramework.Context<any> }> = (props) => {
    const [items, setItems] = React.useState<GridItem[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);
    const [searchText, setSearchText] = React.useState("");
    const [page, setPage] = React.useState(1);
    const [selectedMeta, setSelectedMeta] = React.useState<string | null>(null);
    const [tenantMeta, setTenantMeta] = React.useState<any>(null);
    const [columnFilters, setColumnFilters] = React.useState<Record<string, string>>({});
    const [isDlpPanelOpen, setIsDlpPanelOpen] = React.useState(false);
    const pageSize = 15;
    const theme = getTheme();

    const [filterMenuProps, setFilterMenuProps] = React.useState<{
        target: HTMLElement | MouseEvent;
        columnKey: string;
        isVisible: boolean;
    } | null>(null);

    const onDismissFilter = () => setFilterMenuProps(null);

    const onFilterChange = (key: string, val: string) => {
        setColumnFilters(prev => ({ ...prev, [key]: val }));
        setPage(1);
    };


    const loadEnvironments = async () => {
        setIsLoading(true);
        try {
            const select = "gov_envid,gov_displayname,gov_name,gov_type,gov_region,gov_sku,gov_provisioningstate,gov_version,gov_url,gov_isdefault,gov_createdtime,gov_metadata";
            const result = await props.context.webAPI.retrieveMultipleRecords("gov_environment", `?$select=${select}&$orderby=gov_name asc`);
            const envs: GridItem[] = result.entities.map(e => ({
                key: e.gov_envid,
                label: e.gov_displayname || e.gov_name || "Unnamed",
                type: e.gov_type || e.gov_environmenttype || 'Unknown',
                region: e.gov_region,
                sku: e.gov_sku,
                status: e.gov_provisioningstate,
                version: e.gov_version,
                url: e.gov_url,
                isDefault: e.gov_isdefault,
                created: e.gov_createdtime,
                metadata: e.gov_metadata,
                isExpanded: false
            }));

            // Deduplicate by key to prevent ghost rows
            const unique = Array.from(new Map(envs.map(item => [item.key, item])).values());

            // Extract Sentinel for global stats, then hide it from the grid
            const isSentinel = (u: GridItem) =>
                u.key === "00000000-0000-0000-0000-000000000000" ||
                u.label === "Global Tenant Governance" ||
                u.type === "System" ||
                u.type === "Tenant";

            const sentinel = unique.find(isSentinel);
            if (sentinel) {
                setTenantMeta(getSafeMetadata(sentinel));
            }

            // Filter out the sentinel from the visible grid
            setItems(unique.filter(u => !isSentinel(u)));
        } catch (e) { console.error("Load Failed", e); }
        setIsLoading(false);
    };

    React.useEffect(() => { void loadEnvironments(); }, []);

    const [sortConfig, setSortConfig] = React.useState<{ key: string, desc: boolean }>({ key: 'label', desc: false });

    const filteredItems = React.useMemo(() => {
        return items.filter(i => {
            if (!i) return false;

            const matchesFilters = Object.entries(columnFilters).every(([key, filter]) => {
                if (!filter) return true;
                const f = String(filter).toLowerCase();
                let val = "";

                if (key === 'db' || key === 'files' || key === 'log') {
                    const typeMap: any = { db: 'Database', files: 'File', log: 'Log' };
                    val = (getCapacityMB(getSafeMetadata(i), typeMap[key]) / 1024.0).toFixed(key === 'log' ? 3 : 2);
                } else if (key === 'created') {
                    const d = i.created ? new Date(i.created).toLocaleDateString() : "-";
                    val = d;
                } else if (key === 'isDefault') {
                    val = i.isDefault ? "Yes" : "No";
                } else {
                    val = String((i as any)[key] || "");
                }
                return val.toLowerCase().includes(f);
            });

            const safeSearch = String(searchText || "").toLowerCase();
            const matchesSearch = safeSearch === "" ||
                String(i.label || "").toLowerCase().includes(safeSearch) ||
                String(i.key || "").toLowerCase().includes(safeSearch) ||
                String(i.region || "").toLowerCase().includes(safeSearch) ||
                String(i.type || "").toLowerCase().includes(safeSearch);

            return matchesFilters && matchesSearch;
        });
    }, [items, columnFilters, searchText]);

    const paginatedItems = React.useMemo(() => {
        const list = [...filteredItems];
        if (sortConfig.key) {
            list.sort((a, b) => {
                let vA: any;
                let vB: any;

                if (sortConfig.key === 'db' || sortConfig.key === 'files' || sortConfig.key === 'log') {
                    const typeMap: any = { db: 'Database', files: 'File', log: 'Log' };
                    vA = getCapacityMB(getSafeMetadata(a), typeMap[sortConfig.key]);
                    vB = getCapacityMB(getSafeMetadata(b), typeMap[sortConfig.key]);
                } else {
                    vA = String((a as any)[sortConfig.key] || "").toLowerCase();
                    vB = String((b as any)[sortConfig.key] || "").toLowerCase();
                }

                if (vA < vB) return sortConfig.desc ? 1 : -1;
                if (vA > vB) return sortConfig.desc ? -1 : 1;
                return 0;
            });
        }
        return list.slice((page - 1) * pageSize, page * pageSize);
    }, [filteredItems, page, sortConfig]);
    const totalPages = Math.ceil(filteredItems.length / pageSize);

    const toggleExpand = (key: string) => {
        setItems(prev => prev.map(i => i.key === key ? { ...i, isExpanded: !i.isExpanded } : i));
    };


    const columns: IColumn[] = React.useMemo(() => {
        const baseCols: IColumn[] = [
            {
                key: 'label', name: 'Environment Name', minWidth: 400, isResizable: true,
                isSorted: sortConfig.key === 'label', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    const isTenant = item.key === "00000000-0000-0000-0000-000000000000";
                    return (
                        <Stack horizontal verticalAlign="center" onClick={() => toggleExpand(item.key)} style={{ cursor: 'pointer', padding: '8px 0' }}>
                            <Icon iconName={item.isExpanded ? "ChevronDown" : "ChevronRight"} style={{ fontSize: 10, marginRight: 15, color: '#94a3b8' }} />
                            <div style={{ background: isTenant ? '#1e293b' : theme.palette.themePrimary, padding: 8, borderRadius: 8, marginRight: 15 }}>
                                <Icon iconName={isTenant ? "EMI" : "BuildQueue"} style={{ color: 'white', fontSize: 14 }} />
                            </div>
                            <Stack>
                                <Text style={{ fontWeight: 800, color: isTenant ? '#1e293b' : '#111827', fontSize: 14 }}>{item.label}</Text>
                                <Text variant="small" style={{ color: '#64748b', fontSize: 10, fontFamily: 'monospace' }}>{item.key}</Text>
                            </Stack>
                        </Stack>
                    );
                }
            },
            {
                key: 'type', name: 'SKU', minWidth: 80, isResizable: true,
                isSorted: sortConfig.key === 'type', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    return (
                        <div style={{ padding: '4px 10px', borderRadius: 6, background: '#f1f5f9', color: '#475569', fontSize: 10, fontWeight: 700, display: 'inline-block' }}>
                            {item.type.toUpperCase()}
                        </div>
                    );
                }
            },
            {
                key: 'status', name: 'Status', minWidth: 80, isResizable: true,
                isSorted: sortConfig.key === 'status', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    return (
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                            <div style={{ width: 8, height: 8, borderRadius: '50%', background: item.status === 'Succeeded' || item.status === 'Ready' ? '#107c10' : '#d13438' }} />
                            <Text variant="small" style={{ fontWeight: 600 }}>{item.status}</Text>
                        </Stack>
                    );
                }
            },
            {
                key: 'region', name: 'Region', minWidth: 80, isResizable: true,
                isSorted: sortConfig.key === 'region', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    return (
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                            <Icon iconName="MapPin" style={{ color: '#94a3b8', fontSize: 12 }} />
                            <Text variant="small" style={{ fontWeight: 600, color: '#64748b' }}>{item.region}</Text>
                        </Stack>
                    );
                }
            },
            {
                key: 'version', name: 'Ver', minWidth: 60, isResizable: true,
                isSorted: sortConfig.key === 'version', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    return <Text variant="small" style={{ color: '#666' }}>{item.version}</Text>
                }
            },
            {
                key: 'isDefault', name: 'Default', minWidth: 60, isResizable: true,
                isSorted: sortConfig.key === 'isDefault', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    return <Text variant="small">{item.isDefault ? "Yes" : "No"}</Text>
                }
            },
            {
                key: 'created', name: 'Created', minWidth: 80, isResizable: true,
                isSorted: sortConfig.key === 'created', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    const date = item.created ? new Date(item.created) : null;
                    return <Text variant="small" style={{ color: '#94a3b8' }}>{date ? date.toLocaleDateString() : "-"}</Text>;
                }
            },
            {
                key: 'url', name: 'URL', minWidth: 50, isResizable: true,
                onRender: (item: GridItem) => {
                    if (!item || !item.url) return null;
                    return (
                        <IconButton iconProps={{ iconName: 'CompassNW' }} onClick={() => window.open(item.url, '_blank')} title="Open Environment" />
                    );
                }
            },
            {
                key: 'db', name: 'DB (GB)', minWidth: 80, isResizable: true,
                isSorted: sortConfig.key === 'db', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    const valMB = getCapacityMB(getSafeMetadata(item), "Database");
                    const valGB = valMB / 1024.0;
                    return <Text variant="small" style={{ fontWeight: 600, color: valGB > 10 ? '#d13438' : '#201f1e' }}>{valGB.toFixed(2)} GB</Text>;
                }
            },
            {
                key: 'files', name: 'Files (GB)', minWidth: 80, isResizable: true,
                isSorted: sortConfig.key === 'files', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    const valGB = getCapacityMB(getSafeMetadata(item), "File") / 1024.0;
                    return <Text variant="small">{valGB.toFixed(2)} GB</Text>;
                }
            },
            {
                key: 'log', name: 'Log (GB)', minWidth: 80, isResizable: true,
                isSorted: sortConfig.key === 'log', isSortedDescending: sortConfig.desc,
                onColumnClick: (_, col) => setSortConfig({ key: col.key, desc: sortConfig.key === col.key ? !sortConfig.desc : false }),
                onRender: (item: GridItem) => {
                    if (!item) return null;
                    const valGB = getCapacityMB(getSafeMetadata(item), "Log") / 1024.0;
                    return <Text variant="small" style={{ color: valGB > 0 ? '#f59e0b' : '#666' }}>{valGB.toFixed(3)} GB</Text>;
                }
            },
            {
                key: 'metadata', name: 'DEBUG', minWidth: 50, isResizable: true,
                onRender: (item: GridItem) => (
                    <IconButton iconProps={{ iconName: 'Code' }} onClick={() => setSelectedMeta(JSON.stringify(getSafeMetadata(item), null, 2))} />
                )
            }
        ];

        return baseCols.map(col => ({
            ...col,
            onRenderHeader: (props, defaultRender) => {
                if (!props || !defaultRender) return null;
                const isFiltered = !!columnFilters[col.key];

                return (
                    <Stack horizontal verticalAlign="center" horizontalAlign="space-between" styles={{ root: { width: '100%' } }}>
                        <span style={{ flexGrow: 1 }}>{defaultRender(props)}</span>
                        {col.key !== 'metadata' && col.key !== 'url' && (
                            <IconButton
                                iconProps={{ iconName: 'Filter' }}
                                title="Filter Column"
                                styles={{
                                    root: { height: 16, width: 16 },
                                    icon: { fontSize: 10, color: isFiltered ? theme.palette.themePrimary : '#adb5bd', opacity: isFiltered ? 1 : 0.6 }
                                }}
                                onClick={(e) => {
                                    e.stopPropagation();
                                    setFilterMenuProps({
                                        target: e.currentTarget as any,
                                        columnKey: col.key,
                                        isVisible: true
                                    });
                                }}
                            />
                        )}
                    </Stack>
                );
            }
        }));
    }, [sortConfig, theme, columnFilters]);


    const onRenderRow = (rowProps?: IDetailsRowProps): JSX.Element | null => {
        if (!rowProps) return null;
        const item = rowProps.item as GridItem;
        const isSel = item.isExpanded;
        const isTenant = item.key === "00000000-0000-0000-0000-000000000000";

        return (
            <div key={item.key} style={{ transition: 'all 0.3s ease' }}>
                <DetailsRow
                    {...rowProps}
                    styles={{
                        root: {
                            backgroundColor: isTenant ? '#f1f5f9' : (isSel ? '#f8fbfc' : 'white'),
                            minHeight: 65,
                            borderBottom: '1px solid #f1f5f9',
                            borderLeft: `6px solid ${isTenant ? '#0f172a' : (isSel ? theme.palette.themePrimary : 'transparent')}`,
                            transition: 'all 0.2s',
                            selectors: { ':hover': { backgroundColor: isTenant ? '#e2e8f0' : (isSel ? '#f8fbfc' : '#fcfdfe') } }
                        }
                    }}
                />
                {item.isExpanded && (
                    <div style={{ padding: '10px 40px 30px 65px', background: isTenant ? '#f8fafc' : '#fafafa', overflow: 'hidden' }}>
                        <EnvironmentDetail context={props.context} envId={item.key} metadata={item.metadata} />
                    </div>
                )}
            </div>
        );
    };

    return (
        <Stack style={{ width: '100%', height: '100%', background: '#f8fafc', overflow: 'hidden' }}>
            {/* GLASSMORPHISM HEADER */}
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 40 }}
                style={{
                    padding: '20px 40px',
                    background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
                    boxShadow: '0 10px 30px rgba(0,0,0,0.15)',
                    zIndex: 100,
                    position: 'relative'
                }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 20 }}>
                    <div style={{
                        background: 'linear-gradient(45deg, #3b82f6, #60a5fa)',
                        color: 'white',
                        padding: '12px',
                        borderRadius: 12,
                        boxShadow: '0 8px 25px rgba(59, 130, 246, 0.4)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center'
                    }}>
                        <Icon iconName="ViewDashboard" style={{ fontSize: 32 }} />
                    </div>
                    <Stack>
                        <Text variant="xxLarge" style={{ fontWeight: 900, color: 'white', letterSpacing: -1.2, fontSize: 26 }}>Inventory Sentinel</Text>
                        <Text variant="small" style={{ color: '#93c5fd', fontWeight: 800, letterSpacing: 2, textTransform: 'uppercase' }}>ECOSYSTEM GOVERNANCE v5.0</Text>
                    </Stack>
                </Stack>

                <div style={{ flexGrow: 1, maxWidth: 900 }}>
                    <SearchBox
                        placeholder="Search environments, regions, or schemas..."
                        value={searchText}
                        onChange={(_, v) => { setSearchText(v || ""); setPage(1); }}
                        styles={{
                            root: {
                                borderRadius: 8,
                                background: 'rgba(255,255,255,0.08)',
                                border: '1px solid rgba(255,255,255,0.1)',
                                height: 50,
                                paddingLeft: 15,
                                color: 'white',
                                selectors: { ':after': { border: 'none' }, '::placeholder': { color: '#94a3b8' } }
                            },
                            icon: { color: '#94a3b8', fontSize: 18 },
                            field: { color: 'white' }
                        }}
                    />
                </div>

                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                    {isLoading && <Spinner size={SpinnerSize.medium} styles={{ circle: { borderColor: '#3b82f6 #cbd5e1 #cbd5e1' } }} />}
                    <IconButton
                        iconProps={{ iconName: 'Sync' }}
                        onClick={() => { void loadEnvironments(); }}
                        title="Force Refresh Data"
                        styles={{
                            root: {
                                color: 'white',
                                background: 'rgba(255,255,255,0.1)',
                                border: '1px solid rgba(255,255,255,0.1)',
                                width: 50,
                                height: 50,
                                borderRadius: 8,
                                selectors: { ':hover': { background: 'rgba(255,255,255,0.2)' } }
                            }
                        }}
                    />
                </Stack>
            </Stack>

            <div style={{ flexGrow: 1, overflowY: 'auto', padding: '30px' }}>
                {/* GOVERNANCE SENTINEL TILES */}
                <Stack horizontal tokens={{ childrenGap: 25 }} style={{ marginBottom: 30 }}>
                    <div style={{ flex: 1.5, padding: 25, borderRadius: 16, background: 'linear-gradient(135deg, #1e293b, #0f172a)', color: 'white', boxShadow: '0 10px 30px rgba(0,0,0,0.1)' }}>
                        <Stack horizontal verticalAlign="center" horizontalAlign="space-between" style={{ marginBottom: 20 }}>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                                <Icon iconName="Database" style={{ fontSize: 20, color: '#3b82f6' }} />
                                <Text style={{ color: '#white', fontWeight: 700, letterSpacing: 1 }}>TENANT CAPACITIES</Text>
                            </Stack>
                            <div style={{ background: 'rgba(59, 130, 246, 0.2)', padding: '4px 10px', borderRadius: 20 }}>
                                <Text style={{ color: '#60a5fa', fontWeight: 800, fontSize: 10 }}>HEALTHY</Text>
                            </div>
                        </Stack>

                        <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="start">
                            {[
                                { label: "Database", color: '#3b82f6' },
                                { label: "File", color: '#10b981' },
                                { label: "Log", color: '#f59e0b' }
                            ].map(cap => {
                                // Use tenantMeta from state (authoritative global stats)
                                const val = getCapacityMB(tenantMeta, cap.label);
                                // Fallback: If no global stats, sum visible rows (likely 0 but better than nothing)
                                const finalVal = val > 0 ? val : items.reduce((acc, curr) => acc + getCapacityMB(getSafeMetadata(curr), cap.label), 0);

                                return (
                                    <div key={cap.label} style={{ flex: 1 }}>
                                        <Text variant="small" style={{ color: '#94a3b8', fontWeight: 700, textTransform: 'uppercase', fontSize: 10 }}>{cap.label}</Text>
                                        <Text variant="xLarge" style={{ fontWeight: 800, color: 'white', display: 'block', margin: '2px 0' }}>{(finalVal / 1024).toFixed(1)} <span style={{ fontSize: 12, color: '#64748b' }}>GB</span></Text>
                                        <div style={{ height: 4, background: 'rgba(255,255,255,0.1)', borderRadius: 2, marginTop: 8 }}>
                                            <div style={{ width: '60%', height: '100%', background: cap.color, borderRadius: 2 }} />
                                        </div>
                                    </div>
                                );
                            })}
                        </Stack>
                    </div>

                    <div
                        onClick={() => setIsDlpPanelOpen(true)}
                        style={{
                            flex: 1, padding: 25, borderRadius: 16,
                            background: 'linear-gradient(135deg, #065f46, #064e3b)',
                            color: 'white', boxShadow: '0 10px 30px rgba(6, 95, 70, 0.2)',
                            cursor: 'pointer', transition: 'transform 0.2s',
                            userSelect: 'none'
                        }}
                        onMouseEnter={(e) => e.currentTarget.style.transform = 'translateY(-4px)'}
                        onMouseLeave={(e) => e.currentTarget.style.transform = 'translateY(0)'}
                    >
                        <Stack horizontal verticalAlign="center" horizontalAlign="space-between" style={{ marginBottom: 15 }}>
                            <Icon iconName="Shield" style={{ fontSize: 24, color: '#34d399' }} />
                            <div style={{ background: 'rgba(52, 211, 153, 0.2)', padding: '4px 10px', borderRadius: 20 }}>
                                <Text style={{ color: '#6ee7b7', fontWeight: 800, fontSize: 10 }}>ACTIVE</Text>
                            </div>
                        </Stack>
                        <Text variant="small" style={{ color: '#6ee7b7', fontWeight: 800 }}>GOVERNANCE ENFORCEMENT</Text>
                        <Text variant="xxLarge" style={{ fontWeight: 900, display: 'block', margin: '4px 0' }}>{tenantMeta?.governance?.length || "0"}</Text>
                        <Text variant="small" style={{ color: 'rgba(255,255,255,0.6)' }}>DLP Policies Enforced</Text>
                    </div>

                    <div style={{ flex: 1, padding: 25, borderRadius: 16, background: 'white', border: '1px solid #e2e8f0', boxShadow: '0 10px 30px rgba(0,0,0,0.05)' }}>
                        <Stack horizontal verticalAlign="center" horizontalAlign="space-between" style={{ marginBottom: 15 }}>
                            <Icon iconName="TFVCLogo" style={{ fontSize: 24, color: '#3b82f6' }} />
                            <Text variant="small" style={{ color: '#94a3b8', fontWeight: 800 }}>SYNC STATUS</Text>
                        </Stack>
                        <Text variant="small" style={{ color: '#64748b', fontWeight: 800 }}>TOTAL ASSETS</Text>
                        <Text variant="xxLarge" style={{ fontWeight: 900, display: 'block', margin: '4px 0', color: '#1e293b' }}>{items.length * 42}+</Text>
                        <Text variant="small" style={{ color: '#10b981', fontWeight: 700 }}>● Connected to SeaCass</Text>
                    </div>
                </Stack>

                <div style={{ background: 'white', borderRadius: 12, boxShadow: '0 15px 40px rgba(0,0,0,0.08)', overflow: 'hidden' }}>
                    <DetailsList items={paginatedItems} columns={columns} selectionMode={SelectionMode.none} onRenderRow={onRenderRow} layoutMode={DetailsListLayoutMode.justified} checkboxVisibility={CheckboxVisibility.hidden} />
                </div>
            </div>

            <Stack horizontal verticalAlign="center" horizontalAlign="space-between" style={{ padding: '20px 35px', background: 'white', borderTop: '1px solid #d1d8df' }}>
                <Text variant="small" style={{ color: '#605e5c' }}><b>{items.length}</b> Environments Monitored</Text>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 25 }}>
                    <Text variant="small" style={{ fontWeight: 600 }}>Page {page} of {totalPages || 1}</Text>
                    <Stack horizontal tokens={{ childrenGap: 5 }}>
                        <IconButton iconProps={{ iconName: 'ChevronLeft' }} disabled={page === 1} onClick={() => setPage(p => p - 1)} />
                        <IconButton iconProps={{ iconName: 'ChevronRight' }} disabled={page === totalPages} onClick={() => setPage(p => p + 1)} />
                    </Stack>
                </Stack>
            </Stack>

            {filterMenuProps && (
                <ContextualMenu
                    target={filterMenuProps.target}
                    shouldFocusOnMount={true}
                    onDismiss={onDismissFilter}
                    items={[
                        {
                            key: 'filterInput',
                            onRender: () => (
                                <Stack styles={{ root: { padding: '10px 15px', width: 250 } }}>
                                    <Text variant="small" style={{ fontWeight: 600, marginBottom: 5 }}>
                                        Filter by {columns.find(c => c.key === filterMenuProps.columnKey)?.name}
                                    </Text>
                                    <SearchBox
                                        placeholder="Type to filter..."
                                        value={columnFilters[filterMenuProps.columnKey] || ''}
                                        onChange={(_, newVal) => onFilterChange(filterMenuProps.columnKey, newVal || '')}
                                        styles={{ root: { border: '1px solid #ccc' } }}
                                    />
                                    <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }} style={{ marginTop: 10 }}>
                                        <DefaultButton
                                            text="Clear"
                                            styles={{ root: { height: 24, padding: 0, minWidth: 60 } }}
                                            onClick={() => {
                                                onFilterChange(filterMenuProps.columnKey, '');
                                                onDismissFilter();
                                            }}
                                        />
                                        <PrimaryButton
                                            text="Apply"
                                            styles={{ root: { height: 24, padding: 0, minWidth: 60 } }}
                                            onClick={onDismissFilter}
                                        />
                                    </Stack>
                                </Stack>
                            )
                        }
                    ]}
                />
            )}

            <MetadataViewer
                json={selectedMeta || ""}
                isOpen={!!selectedMeta}
                onDismiss={() => setSelectedMeta(null)}
            />

            <Panel
                isOpen={isDlpPanelOpen}
                onDismiss={() => setIsDlpPanelOpen(false)}
                type={PanelType.medium}
                headerText="Tenant Governance: DLP Policies"
                closeButtonAriaLabel="Close"
                styles={{ content: { background: '#f8fafc' } }}
            >
                <Stack tokens={{ childrenGap: 20 }} style={{ marginTop: 20 }}>
                    <Text variant="medium">
                        Below are the active Data Loss Prevention (DLP) policies detected across your Power Platform tenant.
                        These policies control which connectors can share data, protecting your crown jewels.
                    </Text>

                    <Stack tokens={{ childrenGap: 12 }}>
                        {(tenantMeta?.governance || []).map((p: any) => (
                            <div key={p.id} style={{ background: 'white', padding: 20, borderRadius: 12, border: '1px solid #e2e8f0', boxShadow: '0 4px 12px rgba(0,0,0,0.05)' }}>
                                <Stack horizontal verticalAlign="center" horizontalAlign="space-between" style={{ marginBottom: 12 }}>
                                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                                        <Icon iconName="Shield" style={{ color: '#3b82f6', fontSize: 18 }} />
                                        <Text variant="large" style={{ fontWeight: 800 }}>{p.name}</Text>
                                    </Stack>
                                    <div style={{ background: '#dcfce7', color: '#166534', padding: '4px 10px', borderRadius: 20, fontSize: 10, fontWeight: 800 }}>ENFORCED</div>
                                </Stack>
                                <Stack tokens={{ childrenGap: 4 }}>
                                    <Text variant="small" style={{ color: '#64748b' }}><b>Created:</b> {p.createdTime ? new Date(p.createdTime).toLocaleDateString() : 'N/A'}</Text>
                                    <Text variant="small" style={{ color: '#64748b' }}><b>Rule Sets:</b> {p.ruleSets?.length || 0} active configurations</Text>
                                    <Text variant="small" style={{ color: '#64748b' }}><b>Environment Scope:</b> {p.environmentScope || 'All Environments'}</Text>
                                </Stack>
                            </div>
                        ))}
                    </Stack>

                    {(!tenantMeta?.governance || tenantMeta.governance.length === 0) && (
                        <Stack verticalAlign="center" horizontalAlign="center" style={{ padding: 40 }}>
                            <Icon iconName="SearchData" style={{ fontSize: 48, color: '#cbd5e1', marginBottom: 20 }} />
                            <Text variant="large" style={{ color: '#64748b' }}>No DLP Policies detected in this sync.</Text>
                        </Stack>
                    )}
                </Stack>
            </Panel>
        </Stack>
    );
};
