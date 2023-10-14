/*	Velleman_K8056.ino

  New: Example with SoftwareSerial
  in K8056_Velleman.cpp are the PINs defined: SoftwareSerial mySerial(7, 8); // RX, TX
  Franziska Walter, 10/2023
  
	Programmer		:	A.S. Tilanus (alban.tilanus[at]gmail.com)
	Company			:	reptile-addict.nl
	Version			:	20130421 First release
						20130422 Made some changes as discussed in: http://arduino.cc/forum/index.php/topic,161877.msg1211201.html#msg1211201
	
	Description		:	Example usage of my Velleman K8056 Arduino library.
						http://www.reptile-addict.nl/arduino/?page_id=209

	Keywords		:	Arduino; RS232; Velleman; K8056;

	Copyright		:	http://opensource.org/licenses/mit-license.php

	Disclaimer		:	THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, 	INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

#include <K8056_Velleman.h>

K8056_Velleman Relay;

void setup()
{
	Relay.begin(1); // UNO or NANO can use SoftwareSerial Setting = 1
	Relay.EmergencyStop();
	for (int i = 0; i <= 8; i++) {
		Relay.ON(1,i);
		delay(500);
	}
	delay(500);
	Relay.OFF(1,4);
	delay(2000);
	Relay.EmergencyStop();
	delay(2000);
	Relay.Toggle(1,3);
	delay(1500);
	Relay.Toggle(1,3);
	delay(1500);
	Relay.Toggle(1,3);
	delay(1500);
	Relay.Toggle(1,3);
	delay(2000);
	Relay.digitalWrite(1,6,HIGH); // same as Relay.ON(1,6);
	delay(2000);
	Relay.digitalWrite(1,6,LOW); // same as Relay.OFF(1,6);
}

void loop()
{
	
}
