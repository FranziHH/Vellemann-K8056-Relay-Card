/*	K8056_Velleman.h

	Programmer		:	A.S. Tilanus (alban.tilanus[at]gmail.com)
	Company			:	reptile-addict.nl
	Version			:	20130421 First release
						20130422 Made some changes as discussed in: http://arduino.cc/forum/index.php/topic,161877.msg1211201.html#msg1211201
	
	Description		:	Arduino library to control the Velleman K8056 8-channel relay card thru RS232
						http://www.reptile-addict.nl/arduino/?page_id=209

	Keywords		:	Arduino; RS232; Velleman; K8056;

	Copyright		:	http://opensource.org/licenses/mit-license.php

	Disclaimer		:	THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, 	INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

#ifndef K8056_Velleman_h
	#define K8056_Velleman_h

	#include "Arduino.h"

	class K8056_Velleman
	{
	  public:
		K8056_Velleman();
		void begin();
		void begin(byte SerialAddress);
		void ON(byte CardAddress, byte Relay);
		void OFF(byte CardAddress, byte Relay);
		void EmergencyStop();
		void Toggle(byte CardAddress, byte Relay);
		void digitalWrite(byte CardAddress, byte Relay, byte Status);
		
	  private:
		byte CheckSum(byte Address, byte Instruction, byte Relay);
		void RS232Instruction(byte Address, byte Instruction, byte Relay);
		byte SerialID;
	};
#endif