export interface IDesignerProps {
    imageSrc: string;
    showDesigner: boolean;
    createMiniApp: () => void;
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;
}
