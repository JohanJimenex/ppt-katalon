package ppt

import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testobject.WindowsTestObject as WindowsTestObject
import com.kms.katalon.core.testobject.WindowsTestObject.LocatorStrategy as LocatorStrategy
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows

import internal.GlobalVariable



public abstract class Ppt {

	private enum objects{

		presentacionBlanco("Presentación en blanco"),

		//cambiar color
		tabDiseno("Diseño"),
		btnDarFormatoFondo("Dar formato al fondo"),
		btnColorRelleno("Color de relleno"),
		selectColorNegro("Negro, Texto 1"),
		btnAplicarAtodas("Aplicar a todas"),

		tabInsertar("Insertar"),
		btnVideo("Vídeo"),
		selectVideoEnMiPC("Vídeo en Mi PC..."),
		btnInsertar('Insertar'),

		//Tab Reproduccion
		tabReproduccion("Reproducción"),
		checkRepetir("Repetir la reproducción hasta su interrupción"),
		checkRebobinar("Rebobinar después de la reproducción"),
		checkReproducir("Reproducir a pantalla completa"),
		btnNuevaDiapositiva("Nueva diapositiva"),
		btnEnBlanco("En blanco"),

		//Guardar
		btnCancelar("Cancelar"),
		btnGuardar("Guardar"),
		btnEscritorio("Escritorio");

		public WindowsTestObject obj

		private objects(String value){
			this.obj = new WindowsTestObject('Name')
			this.obj.setLocatorStrategy(LocatorStrategy.NAME)
			this.obj.setLocator(value)
		}
	}



	public static void agregarVideos(int cantidad){

		Windows.startApplicationWithTitle('C:\\\\Program Files\\\\Microsoft Office\\\\root\\\\Office16\\\\POWERPNT.exe', 'PowerPoint')

		Windows.click(objects.presentacionBlanco.obj)

		//Cambiar color de fondo
		Windows.click(objects.tabDiseno.obj)
		Windows.click(objects.btnDarFormatoFondo.obj)
		Windows.click(objects.btnColorRelleno.obj)
		Windows.click(objects.selectColorNegro.obj)
		Windows.click(objects.btnAplicarAtodas.obj)

		Windows.click(objects.tabInsertar.obj)

		for (int i = 1; i < cantidad; i++) {

			Windows.waitForElementPresent(objects.btnVideo.obj, GlobalVariable.tiempoEspera)
			Windows.click(objects.btnVideo.obj)
			Windows.click(objects.selectVideoEnMiPC.obj)

			WindowsTestObject video = new WindowsTestObject('video')
			video.setLocatorStrategy(LocatorStrategy.NAME)
			video.setLocator('m ('+i+').avi')

			Windows.waitForElementPresent(video, GlobalVariable.tiempoEspera)
			Windows.click(video)

			Windows.click(objects.btnInsertar.obj)

			Windows.waitForElementPresent(objects.tabReproduccion.obj, GlobalVariable.tiempoEspera)
			Windows.click(objects.tabReproduccion.obj)

			Windows.waitForElementPresent(objects.checkRepetir.obj, GlobalVariable.tiempoEspera)
			Windows.click(objects.checkRepetir.obj)
			Windows.click(objects.checkRebobinar.obj)
			Windows.click(objects.checkReproducir.obj)

			//		Windows.sendKeys(findWindowsObject('btnInsertar'), Keys.chord(Keys.CONTROL, 'm'))
			Windows.click(objects.tabInsertar.obj)

			Windows.switchToDesktop()
			Windows.switchToApplication()

			Windows.waitForElementPresent(objects.btnNuevaDiapositiva.obj, GlobalVariable.tiempoEspera)
			Windows.click(objects.btnNuevaDiapositiva.obj)

		}
	}
}

