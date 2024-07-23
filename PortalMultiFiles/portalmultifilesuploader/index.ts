import { IInputs, IOutputs } from "./generated/ManifestTypes";
class EntityReference {
  constructor(public typeName: string, public id: string) { }
}

class AttachedFile implements ComponentFramework.FileObject {
  constructor(
    public annotationId: string,
    public fileName: string,
    public mimeType: string,
    public fileContent: string,
    public fileSize: number
  ) { }
}

interface FileNode {
  id: string;
  file: File;
}

export class portalmultifilesuploader
  implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private Files: FileNode[] = [];
  private entityReference: EntityReference;
  private _context: ComponentFramework.Context<IInputs>;

  constructor() { }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ) {
    this.Files = [];
    this._context = context;
    this.entityReference = new EntityReference(
      (<any>context).page.entityTypeName,
      (<any>context).page.entityId
    );
    const UploadForm = this.CreateFormUploadDiv();
    container.appendChild(UploadForm);
  }

  private CreateFormUploadDiv = (): HTMLDivElement => {
    const UploadForm = document.createElement("div");
    const UploadLabel = document.createElement("label");
    UploadLabel.htmlFor = "file-upload";
    UploadLabel.id = "lbl-file-upload";
    UploadLabel.innerText = "Choose Files to Upload";
    const UploadInput = document.createElement("input");
    UploadInput.id = "file-upload";
    UploadInput.type = "file";
    UploadInput.multiple = true;
    UploadInput.addEventListener("change", this.handleBrowse);
    const DragDiv = document.createElement("Div");
    DragDiv.id = "watermarkdiv";
    DragDiv.className = "watermarkdiv";
    DragDiv.innerText = "or drop files here...";

    const catchedfileslist = document.createElement("ol");
    catchedfileslist.id = "catchedfileslist";
    const fileCatcher = this.createDiv("files-catcher", "files", [
      catchedfileslist,
    ]);

    const filesHolder = this.createDiv("file-holder", "", [
      DragDiv,
      fileCatcher,
    ]);

    const UploadButton = document.createElement("button");
    UploadButton.innerText = "Upload";
    UploadButton.className = "buttons";
    UploadButton.addEventListener("click", this.handleUpload);
    const ClearButton = document.createElement("button");
    ClearButton.innerText = "Reset";
    ClearButton.className = "buttons";
    ClearButton.addEventListener("click", this.handleReset);
    const leftDiv = this.createDiv("left-container", "left-container", [
      UploadLabel,
      UploadInput,
      UploadButton,
      ClearButton,
    ]);

    const rightDiv = this.createDiv("right-container", "right-container", [
      filesHolder,
    ]);
    rightDiv.addEventListener("dragover", this.FileDragHover);
    rightDiv.addEventListener("dragleave", this.FileDragHover);
    rightDiv.addEventListener("drop", this.handleBrowse);
    const mainContainer = this.createDiv("main-container", "main-container", [
      leftDiv,
      rightDiv,
    ]);
    UploadForm.appendChild(mainContainer);

    return UploadForm;
  };

  private createDiv(
    divid: string,
    classname = "",
    childElements?: HTMLElement[]
  ): HTMLDivElement {
    const _div: HTMLDivElement = document.createElement("div");
    _div.id = divid;
    classname ? (_div.className = classname) : "";
    if (childElements != null && childElements?.length > 0) {
      childElements.forEach((child) => {
        _div.appendChild(child);
      });
    }
    return _div;
  }

  private handleBrowse = (e: any): void => {
    e.preventDefault();
    console.log("handleBrowse");
    console.log(e);
    const files = e.target.files || e.dataTransfer.files;
    if (files.length > 0) {
      this.addFiles(files);
    }
  };

  addFiles(files: FileList) {
    let counter = this.Files.length;
    if (counter > 0 || files.length > 0) {
      const filesDiv = this.$id("watermarkdiv") as HTMLDivElement;
      filesDiv.style.display = "none";
    }
    const fileList = this.$id("catchedfileslist") as HTMLOListElement;
    for (let i = 0; i < files.length; i++) {
      counter++;
      const nodetype = {} as FileNode;
      nodetype.id = "progress" + counter;
      nodetype.file = files[i];
      const fileNode = document.createElement("li");
      fileNode.className = "fileNode";
      const text = document.createTextNode(files[i].name);
      fileNode.appendChild(text);
      fileList.appendChild(fileNode);
      this.Files[this.Files.length] = nodetype;
    }
  }

  private handleUpload = async (e: any): Promise<void> => {
    console.log("handleUpload");
    console.log(e);
    const files = this.Files;
    const valid = files && files.length > 0;
    if (!valid) {
      alert("Please select a file!");
      return;
    }

    const uploadedFiles: AttachedFile[] = [];

    for (let i = 0; i < files.length; i++) {
      const file = files ? files[i].file : "";
      if (file !== "") {
        await new Promise<void>((resolve) => {
          this.toBase64String(file, (file: File, text: string) => {
            const type = file.type;
            const notesEntity = new AttachedFile(
              "",
              file.name,
              type,
              text,
              file.size
            );
            uploadedFiles.push(notesEntity);
            resolve();
          });
        });
      }
    }

    if (uploadedFiles.length > 0) {
      await this.addAttachments(uploadedFiles);
    }

    alert(`uploaded ${files.length} number of files as attachments`);
    this.clearAttachments();
  };

  clearAttachments = (): void => {
    const fileList = document.getElementById("catchedfileslist") as any;
    if (fileList) {
      while (fileList.hasChildNodes()) {
        fileList.removeChild(fileList.firstChild);
      }
    }
    this.Files = [];
    this.$id("watermarkdiv").style.display = "block";
  };

  addAttachments = async (files: AttachedFile[]): Promise<void> => {
    const fileData = files.map((file) => {
      const notesEntity: any = {};
      let fileContent = file.fileContent;
      fileContent = fileContent.substring(
        fileContent.indexOf(",") + 1,
        fileContent.length
      );
      notesEntity["documentbody"] = fileContent;
      notesEntity["filename"] = file.fileName;
      notesEntity["filesize"] = file.fileSize;
      notesEntity["mimetype"] = file.mimeType;
      notesEntity["subject"] = file.fileName;
      notesEntity["objecttypecode"] = this.entityReference.typeName;
      notesEntity[
        `objectid_${this.entityReference.typeName}@odata.bind`
      ] = `/${this.CollectionNameFromLogicalName(
        this.entityReference.typeName
      )}(${this.entityReference.id})`;

      return {
        index: `1`,
        base64: notesEntity.documentbody,
        mimeType: notesEntity.mimetype,
        fileName: notesEntity.filename,
        mainbase64: "data:" + notesEntity.mimetype + ";base64," + notesEntity.documentbody,
      };
    });

    // @ts-expect-error: Xrm may not be present in window.parent in some environments
    if (window.parent.Xrm == undefined) {
      const url = 'https://prod2-07.centralindia.logic.azure.com:443/workflows/45477bdfc6ef4ce6944f6263af1d4c42/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7cKgxerbqoAQsrxhD2PlqrW9TgR-AG_x8WJb-ZmN9aw';

      const myApiResult = await fetch(url, {
        method: 'POST', // or 'PUT'
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(fileData),
      });

      const result = await myApiResult.json();

      if (myApiResult.status == 200) {
        alert('Uploaded Successfully');
      } else {
        alert('Please contact admin');
      }
    }
  };

  CollectionNameFromLogicalName = (entityLogicalName: string): string => {
    if (entityLogicalName[entityLogicalName.length - 1] != "s") {
      return `${entityLogicalName}s`;
    } else {
      return `${entityLogicalName}ies`;
    }
  };

  private toBase64String = (
    file: File,
    successFn: (file: File, body: string) => void
  ) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => successFn(file, reader.result as string);
    console.log(reader.result);
    return reader.result;
  };

  private $id = (id: string): any => {
    return document.getElementById(id);
  };

  private handleReset = (e: any): void => {
    console.log("handleReset");
    console.log(e);
    this.clearAttachments();
  };

  private FileDragHover = (e: any): void => {
    e.stopPropagation();
    e.preventDefault();
    console.log("dragover", e);
  };

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Add code to update control view
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void { }
}
