declare module "bpmn-js/lib/Modeler" {
  type ConstructorOptions = {
    container: HTMLElement;
    keyboard?: {
      bindTo?: Window;
    };
  };

  export default class BpmnModeler {
    constructor(options: ConstructorOptions);
    destroy(): void;
    get<T = unknown>(serviceName: string): T;
    importXML(xml: string): Promise<{ warnings: unknown[] }>;
    saveXML(options?: { format?: boolean }): Promise<{ xml?: string }>;
  }
}

declare module "bpmn-js/lib/NavigatedViewer" {
  type ConstructorOptions = {
    container: HTMLElement;
  };

  export default class BpmnViewer {
    constructor(options: ConstructorOptions);
    destroy(): void;
    get<T = unknown>(serviceName: string): T;
    importXML(xml: string): Promise<{ warnings: unknown[] }>;
    saveXML(options?: { format?: boolean }): Promise<{ xml?: string }>;
  }
}

