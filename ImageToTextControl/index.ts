import {IInputs, IOutputs} from "./generated/ManifestTypes";

///<reference path="./node_modules/jquery/jquery.min.js" />

import * as $ from "jquery";

const ShowErrorClassName = "ShowError";

export class ImageToTextControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// PCF framework context, "Input Properties" containing the parameters, control metadata and interface functions.
	private _context: ComponentFramework.Context<IInputs>;

	// PCF framework delegate which will be assigned to this object which would be called whenever any update happens.
	private _notifyOutputChanged: () => void;

	private controlContainer: HTMLDivElement;
	private cameraButton: HTMLButtonElement;
	private uploadedImage: HTMLImageElement;
	private textResult : HTMLLabelElement;
	private statusText : HTMLLabelElement;

	//Update with your Subscription Key and the EndPoint url
	private subscriptionKey = "<subscription key goes here>";
	private endpoint = "<computer vision endpoint goes here>";

	/**
	 * Empty constructor.
	 */
	constructor()
	{}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Add control initialization code
		this._context = context;

		this.controlContainer = document.createElement("div");

		//Creating an input element to get the question.
		this.textResult = document.createElement("label");
		this.textResult.classList.add("text_result_style");

		this.uploadedImage = document.createElement("img");		
		this.uploadedImage.classList.add("uploaded_image_style");

		//Create an upload button to call the forms recognizer api
		this.cameraButton = document.createElement("button");
		this.cameraButton.className = "fas fa-camera fa-2x camera_button_style";
		this.cameraButton.addEventListener("click", this.onCameraButtonClick.bind(this));

		//Creating a label element to display the answer.
		this.statusText = document.createElement("label");
		this.statusText.setAttribute("type", "label");
		this.statusText.classList.add("answer_Input_Style");
		
		// Adding the label and button created to the container DIV.
		this.controlContainer.className = "control_container";
		this.controlContainer.appendChild(this.cameraButton);
		this.controlContainer.appendChild(this.uploadedImage);
		this.controlContainer.appendChild(this.textResult);
		this.controlContainer.appendChild(this.statusText);
		container.appendChild(this.controlContainer);

		this._notifyOutputChanged = notifyOutputChanged;
	}

	private showError(errorText : string): void
	{
		this.enableCameraButton();
		this.statusText.innerHTML = errorText;
		this.statusText.classList.add(ShowErrorClassName);
	}

	private hideError(): void
	{
		this.enableCameraButton();	
		this.statusText.innerHTML = "";
		this.statusText.classList.remove(ShowErrorClassName);
	}

	private enableCameraButton(): void
	{
		this.cameraButton.removeAttribute("disabled");	
	}

	private disableCameraButton(): void
	{	
		this.cameraButton.setAttribute("disabled", "disabled");			
	}

	private onCameraButtonClick(event: Event): void
	{
		//Invoke Device capture api 
		this._context.device.captureImage().then(this.processFile.bind(this), this.showError.bind(this));
	}

	private generateImageSrcUrl(fileType: string, fileContent: string): string
	{
		return  "data:image/" + fileType + ";base64, " + fileContent;
	}		
	
	//This will transpose the base64 url to binary format. 
	//The Cognitive Service api expects the image in the raw binary format
	private makeBlob(dataURL : string) 
	{
		let BASE64_MARKER = ';base64,';
		if (dataURL.indexOf(BASE64_MARKER) == -1) {
			let parts = dataURL.split(',');
			let contentType = parts[0].split(':')[1];
			let raw = decodeURIComponent(parts[1]);
			return new Blob([raw], { type: contentType });
		}
		let parts = dataURL.split(BASE64_MARKER);
		let contentType = parts[0].split(':')[1];
		let raw = window.atob(parts[1]);
		let rawLength = raw.length;

		let uInt8Array = new Uint8Array(rawLength);

		for (let i = 0; i < rawLength; ++i) {
			uInt8Array[i] = raw.charCodeAt(i);
		}
		return new Blob([uInt8Array], { type: contentType });
	}

	private processFile(file: ComponentFramework.FileObject): void
	{		
		this.statusText.innerHTML = this._context.resources.getString("PCF_ImageToTextControl_ImageProcessing_Message"); 
		//Disable and change the Button text to indicate the camera control is initiated and text parsing is in progress.
		this.disableCameraButton();	

		let uriBase = this.endpoint + "vision/v2.0/read/core/asyncBatchAnalyze";
		let fileExtension: string = '';
		let imageUrl:string = '';

		try {
			if (file && file.fileName) {			
				fileExtension = file.fileName.split('.').pop() || '';
				if ((fileExtension) && ((fileExtension == "jpeg") || (fileExtension == "PNG") || (fileExtension == "png")|| (fileExtension == "jpg"))) {					
					let dataURL = this.generateImageSrcUrl(fileExtension, file.fileContent);					
					imageUrl = dataURL;
					this.uploadedImage.setAttribute("src", imageUrl);
				}
				else {		
					//If the file type is incorrect, do not proceed.		
					this.showError(this._context.resources.getString("PCF_ImageToTextControl_ImageType_NotSupported_Error"));
					return;
				}
			}
			else
			{
				this.hideError();
				//The user cancelled the camera process, do not proceed.				
				return;
			}		
		}
		catch (err) {
			this.showError(err);
		}

		$.ajax({
            url: uriBase,            
            beforeSend: (jqXHR) => {
                jqXHR.setRequestHeader("Content-Type","application/octet-stream");
                jqXHR.setRequestHeader("Ocp-Apim-Subscription-Key", this.subscriptionKey);
			},
			mimeType: "application/octet-stream",
			processData: false,
			type: "POST",
			contentType: false,           		
			data: this.makeBlob(imageUrl),
			success: (data, textStatus, jqXHR) => {

			// Show progress.	
			//"Waiting 3 seconds to retrieve the recognized text.");
			this.statusText.innerHTML = this._context.resources.getString("PCF_ImageToTextControl_TextSubmitted_Message");
			this._notifyOutputChanged();

			// Note: The response may not be immediately available. Text
			// recognition is an asynchronous operation that can take a variable
			// amount of time depending on the length of the text you want to
			// recognize. You may need to wait or retry the GET operation.
			
			//You may need to wait for more seconds if the document to be parsed is larger
			setTimeout( () => {
			// "Operation-Location" in the response contains the URI to retrieve the recognized text.
			let operationLocation = jqXHR.getResponseHeader("Operation-Location") || '';

			// Make the second REST API call and get the response.
				$.ajax({
				url: operationLocation,
				beforeSend: (jqXHR) => {
					jqXHR.setRequestHeader("Content-Type","application/json");
					jqXHR.setRequestHeader(
						"Ocp-Apim-Subscription-Key", this.subscriptionKey);
				},
				type: "GET", 
				success: (data) => {					
				let extractedResponseValue : string = "";

				//Stitch the responses together
				$.each(data.recognitionResults[0].lines, (index, value) =>
				{
					extractedResponseValue = extractedResponseValue + " " + value.text;
				});

				//If the uploaded image has some text to be parsed
				if(extractedResponseValue != "")
				{
					this.textResult.innerHTML = extractedResponseValue;
					this.hideError();
				}
				//If the uploaded image does not contain any text to be parsed
				else
				{
					this.textResult.innerHTML = "";
					this.showError(this._context.resources.getString("PCF_ImageToTextControl_NothingToBeParsed_Message"));
				}
				this._notifyOutputChanged();
				},
				error: (jqXHR, textStatus, errorThrown) => {					
					this.parseError(jqXHR, errorThrown);
				}
			})			
			}, 3000);
			},
			error: (jqXHR, textStatus, errorThrown) => {
				this.parseError(jqXHR, errorThrown);
				}
			});					
	}

	private parseError(jqXHR: JQuery.jqXHR, errorThrown: string)
	{
		var errorString = (errorThrown === "") ? "Error. " :
		errorThrown + " (" + jqXHR.status + "): ";
		errorString += (jqXHR.responseText === "") ? "" :
		($.parseJSON(jqXHR.responseText).message) ?
		$.parseJSON(jqXHR.responseText).message :
		$.parseJSON(jqXHR.responseText).error.message;				
		this.showError(errorString);
		this._notifyOutputChanged();	
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
	}

	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
		this.cameraButton.removeEventListener("click", this.onCameraButtonClick.bind(this));
	}
}