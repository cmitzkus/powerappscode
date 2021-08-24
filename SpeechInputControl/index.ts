import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { ResultReason } from "microsoft-cognitiveservices-speech-sdk";
import axios from "axios";

import * as speechsdk from "microsoft-cognitiveservices-speech-sdk";
import { debug } from "console";

export class SpeechInputControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	private _container: HTMLDivElement;
	private _audioContext: AudioContext;
	private _context: ComponentFramework.Context<IInputs>;
	private _notifyOutputChanged: () => void;
	private _refreshData: EventListenerOrEventListenerObject;
	private recognizedText: string;
	private speechLanguage: string;
	private speechKey: string;
	private speechZone: string;
	private working: string;

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(
		context: ComponentFramework.Context<IInputs>, 
		notifyOutputChanged: () => void, 
		state: ComponentFramework.Dictionary, 
		container:HTMLDivElement): void
	{
		this._audioContext = new AudioContext();
		this._context = context;
		this._container = document.createElement("div");
		this._notifyOutputChanged = notifyOutputChanged;
		this._audioContext.resume().then(() => {
			console.log('Audio resumed!');
		});

		// retrieving the latest value from the control and setting it to the HTMl elements.
		this.recognizedText = context.parameters.recognizedText.raw!;
		this.speechLanguage = context.parameters.speechLanguage.raw!;
		this.speechKey = context.parameters.speechKey.raw!;
		this.speechZone = context.parameters.speechZone.raw!;
		this.working = context.parameters.working.raw!;

		// appending the HTML elements to the control's HTML container element.
		container.appendChild(this._container);

		this.sttFromMic();
	}

	public async getTokenOrRefresh() {
		const speechKey = this.speechKey;
		const speechRegion = this.speechZone;
	
		const headers = {
		  headers: {
			"Ocp-Apim-Subscription-Key": speechKey,
			"Content-Type": "application/x-www-form-urlencoded",
		  },
		};
	
		try {
		  const tokenResponse = await axios.post(
			`https://${speechRegion}.api.cognitive.microsoft.com/sts/v1.0/issueToken`,
			null,
			headers
		  );
		  console.log("Token fetched from back-end: " + tokenResponse.data);
		  return { authToken: tokenResponse.data, region: speechRegion };
		} catch (err) {
		  return { authToken: null, error: err.response.data };
		}
	}

	public async sttFromMic() {

		const tokenObj = await this.getTokenOrRefresh();

		const speechConfig = speechsdk.SpeechConfig.fromAuthorizationToken(
		  tokenObj.authToken,
		  tokenObj.region || ""
		);

		speechConfig.speechRecognitionLanguage = this.speechLanguage; //"en-US"
	
		const audioConfig = speechsdk.AudioConfig.fromDefaultMicrophoneInput();
		const recognizer = new speechsdk.SpeechRecognizer(
		  speechConfig,
		  audioConfig
		);
	
		this.recognizedText = "...";
	
		recognizer.recognizeOnceAsync((result) => {
			let recognizedText;
			if (result.reason === ResultReason.RecognizedSpeech) {
				recognizedText = result.text;
				console.log(recognizedText);
				this.working = "done";
			} else {
				recognizedText =
				"error";
			}
			this.recognizedText = recognizedText;
		  	this._notifyOutputChanged();
		});
	}

	/** 
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
		this.recognizedText = context.parameters.recognizedText.raw!;
		this.speechLanguage = context.parameters.speechLanguage.raw!;
		this.speechKey = context.parameters.speechKey.raw!;
		this.speechZone = context.parameters.speechZone.raw!;
		this.working = context.parameters.working.raw!;
		this._notifyOutputChanged();
	}

	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		if(this.working == "start"){
			this.sttFromMic();
			this.working = "recording";
		}
		return {
			recognizedText : this.recognizedText,
			speechLanguage : this.speechLanguage,
			speechKey : this.speechKey,
			speechZone : this.speechZone,
			working : this.working
		};
	}

	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}
}
