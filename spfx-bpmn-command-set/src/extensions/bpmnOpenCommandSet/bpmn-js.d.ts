declare module 'bpmn-js/lib/Modeler' {
  export default class BpmnModeler {
    public constructor(options: { container: HTMLElement; keyboard?: { bindTo: Window } });
    public destroy(): void;
    public get<T = unknown>(name: string): T;
    public importXML(xml: string): Promise<{ warnings?: Array<Error | { message?: string }> }>;
    public saveXML(options?: { format?: boolean }): Promise<{ xml?: string }>;
  }
}

declare module 'bpmn-js/lib/NavigatedViewer' {
  export default class BpmnViewer {
    public constructor(options: { container: HTMLElement });
    public destroy(): void;
    public get<T = unknown>(name: string): T;
    public importXML(xml: string): Promise<{ warnings?: Array<Error | { message?: string }> }>;
    public saveXML(options?: { format?: boolean }): Promise<{ xml?: string }>;
  }
}

declare module '*.css';
