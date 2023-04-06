#define _WIN32_DCOM
#define _CRT_SECURE_NO_WARNINGS

#include <windows.h>
#include <pathcch.h>
#include <powrprof.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <wincred.h>
#include <format>
#include <string>
#include <chrono>
#include <vector>
#include <fstream>
#include <taskschd.h>

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")
#pragma comment(lib, "Pathcch.lib")
#pragma comment(lib, "PowrProf.lib")

using namespace std;

#define ERROR_THROW(message) throw exception(format(message##": {:#x}",static_cast<unsigned long>(hr)).c_str())
#define IF_ERROR_THROW(message) if (FAILED(hr)) ERROR_THROW(message)

#define ERROR_THROWF(message, ...) throw exception(format("{}: {:#x}",format(message,__VA_ARGS__),static_cast<unsigned long>(hr)).c_str())
#define IF_ERROR_THROWF(message, ...) if (FAILED(hr)) ERROR_THROWF(message,__VA_ARGS__)

#define LAZY_STR(wstr) ((const char*)(wstr.c_str()))

struct Task {
	ITaskDefinition* pTask = NULL;
	ITaskSettings* pSettings = NULL;
	IRegistrationInfo* pRegInfo = NULL;
	IPrincipal* pPrincipal = NULL;
	IIdleSettings* pIdleSettings = NULL;
	ITriggerCollection* pTriggerCollection = NULL;
	IActionCollection* pActionCollection = NULL;
	wstring taskName;
	HRESULT& hr;

	Task(const wstring& _taskName, ITaskDefinition* _pTask, HRESULT& _hr) : taskName(_taskName), pTask(_pTask), hr(_hr) {
		hr = pTask->get_Settings(&pSettings);
		if (FAILED(hr))
		{
			pTask->Release();
			pTask = NULL;
			ERROR_THROW("Cannot get settings pointer");
		}
	}

	~Task() {
		if (pTriggerCollection != NULL) pTriggerCollection->Release();
		if (pActionCollection != NULL) pActionCollection->Release();
		if (pIdleSettings != NULL) pIdleSettings->Release();
		if (pPrincipal != NULL) pPrincipal->Release();
		if (pSettings != NULL) pSettings->Release();
		if (pRegInfo != NULL) pRegInfo->Release();
		if (pTask != NULL) pTask->Release();

		pTriggerCollection = NULL;
		pActionCollection = NULL;
		pIdleSettings = NULL;
		pPrincipal = NULL;
		pSettings = NULL;
		pRegInfo = NULL;
		pTask = NULL;
	}

	void SetAuthor(const wstring& author) {
		if (pRegInfo == NULL) {
			hr = pTask->get_RegistrationInfo(&pRegInfo);
			IF_ERROR_THROW("Cannot get identification pointer");
		}

		hr = pRegInfo->put_Author(_bstr_t(author.c_str()));
		IF_ERROR_THROWF("Cannot put identification info ({})", LAZY_STR(author));
	}

	void SetLogonType(TASK_LOGON_TYPE logonType = TASK_LOGON_INTERACTIVE_TOKEN, TASK_RUNLEVEL_TYPE runLevel = TASK_RUNLEVEL_HIGHEST) {
		if (pPrincipal == NULL) {
			hr = pTask->get_Principal(&pPrincipal);
			IF_ERROR_THROW("Cannot get principal pointer");
		}

		hr = pPrincipal->put_Id(_bstr_t(L"Principal1"));
		IF_ERROR_THROW("Cannot put the principal ID");

		hr = pPrincipal->put_LogonType(logonType);
		IF_ERROR_THROWF("Cannot put logon type ({})", (int)logonType);

		hr = pPrincipal->put_RunLevel(runLevel);
		IF_ERROR_THROWF("Cannot set run level ({})", (int)runLevel);
	}

	void SetStartWhenAvailable(VARIANT_BOOL b = VARIANT_TRUE) {
		hr = pSettings->put_StartWhenAvailable(b);
		IF_ERROR_THROW("Cannot set 'start when available'");
	}

	void SetStopOnBatteries(VARIANT_BOOL b = VARIANT_FALSE) {
		hr = pSettings->put_DisallowStartIfOnBatteries(b);
		IF_ERROR_THROW("Cannot set 'disallow start if on batteries'");

		hr = pSettings->put_StopIfGoingOnBatteries(b);
		IF_ERROR_THROW("Cannot set 'stop if going on batteries'");
	}

	void SetTimeLimit(const wstring& limit = L"PT0S") {
		hr = pSettings->put_ExecutionTimeLimit(_bstr_t(limit.c_str()));
		IF_ERROR_THROWF("Cannot set time limit ({})", LAZY_STR(limit));
	}

	void SetIdleSettings(const wstring& timeoutSettings = L"PT5M", VARIANT_BOOL stopOnIdleEnd = VARIANT_FALSE) {
		if (pIdleSettings == NULL) {
			hr = pSettings->get_IdleSettings(&pIdleSettings);
			IF_ERROR_THROW("Cannot get idle setting information");
		}

		hr = pIdleSettings->put_WaitTimeout(_bstr_t(timeoutSettings.c_str()));
		IF_ERROR_THROWF("Cannot put wait time out ({})", LAZY_STR(timeoutSettings));

		hr = pIdleSettings->put_StopOnIdleEnd(stopOnIdleEnd);
		IF_ERROR_THROW("Cannot put stop on idle end");
	}

	ITimeTrigger* AddTimeTrigger(const wstring& start) {
		if (pTriggerCollection == NULL) {
			hr = pTask->get_Triggers(&pTriggerCollection);
			IF_ERROR_THROW("Cannot get trigger collection");
		}

		long triggerCount;
		hr = pTriggerCollection->get_Count(&triggerCount);
		IF_ERROR_THROW("Cannot get number of triggers");

		ITrigger* pTrigger = NULL;
		hr = pTriggerCollection->Create(TASK_TRIGGER_TIME, &pTrigger);
		IF_ERROR_THROW("Cannot create trigger");

		ITimeTrigger* pTimeTrigger = NULL;
		hr = pTrigger->QueryInterface(IID_ITimeTrigger, (void**)&pTimeTrigger);
		pTrigger->Release();
		IF_ERROR_THROW("QueryInterface call failed for ITimeTrigger");

		wstring ID = format(L"Trigger{}", triggerCount + 1);
		hr = pTimeTrigger->put_Id(_bstr_t(ID.c_str()));
		if (FAILED(hr)) {
			pTimeTrigger->Release();
			ERROR_THROWF("Cannot put trigger ID ({})", LAZY_STR(ID));
		}

		hr = pTimeTrigger->put_StartBoundary(_bstr_t(start.c_str()));
		if (FAILED(hr)) {
			pTimeTrigger->Release();
			ERROR_THROWF("Cannot add start boundary to trigger ({})", LAZY_STR(start));
		}
		
		return pTimeTrigger;
	}

	ITimeTrigger* AddTimeTrigger(const wstring& start, const wstring& end) {
		ITimeTrigger* pTimeTrigger = AddTimeTrigger(start);

		hr = pTimeTrigger->put_EndBoundary(_bstr_t(end.c_str()));
		if (FAILED(hr)) {
			pTimeTrigger->Release();
			ERROR_THROWF("Cannot put end boundary on trigger ({})", LAZY_STR(end));
		}

		return pTimeTrigger;
	}

	IExecAction* AddExecutableAction(const wstring& execPath) {
		if (pActionCollection == NULL) {
			hr = pTask->get_Actions(&pActionCollection);
			IF_ERROR_THROW("Cannot get Task collection pointer");
		}

		IAction* pAction = NULL;
		hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
		IF_ERROR_THROW("Cannot create the action");

		IExecAction* pExecAction = NULL;
		hr = pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);
		pAction->Release();
		IF_ERROR_THROW("QueryInterface call failed for IExecAction");

		hr = pExecAction->put_Path(_bstr_t(execPath.c_str()));
		if (FAILED(hr)) {
			pExecAction->Release();
			ERROR_THROWF("Cannot put action path ({})", LAZY_STR(execPath));
		}

		return pExecAction;
	}

	IExecAction* AddExecutableAction(const wstring& execPath, const wstring& folderPath) {
		IExecAction* pExecAction = AddExecutableAction(execPath);

		hr = pExecAction->put_WorkingDirectory(_bstr_t(folderPath.c_str()));
		if (FAILED(hr)) {
			pExecAction->Release();
			ERROR_THROWF("Cannot put action working directory ({})", LAZY_STR(folderPath));
		}

		return pExecAction;
	}
};

wstring HResultToString(HRESULT hr) {
	DWORD errorMessageID = hr == S_OK ? ERROR_SUCCESS : ((hr & 0xFFFF0000) == MAKE_HRESULT(SEVERITY_ERROR, FACILITY_WIN32, 0) ? HRESULT_CODE(hr) : ERROR_CAN_NOT_COMPLETE);

	LPWSTR messageBuffer = nullptr;

	size_t size = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL, errorMessageID, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&messageBuffer, 0, NULL);

	wstring message(messageBuffer, size);

	LocalFree(messageBuffer);

	return message;
}

class TaskService {
private:
	bool _initialised = false;
	HRESULT hr;
	ITaskService* pService = NULL;
	ITaskFolder* pRootFolder = NULL;

public:
	TaskService() {
		if (_initialised) return;

		hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
		IF_ERROR_THROW("CoInitializeEx failed");

		hr = CoInitializeSecurity(
			NULL,
			-1,
			NULL,
			NULL,
			RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
			RPC_C_IMP_LEVEL_DELEGATE,
			NULL,
			0,
			NULL);

		if (FAILED(hr))
		{
			CoUninitialize();
			ERROR_THROW("CoInitializeSecurity failed");
		}

		hr = CoCreateInstance(CLSID_TaskScheduler,
							   NULL,
							   CLSCTX_INPROC_SERVER,
							   IID_ITaskService,
							   (void**)&pService);

		if (FAILED(hr))
		{
			CoUninitialize();
			ERROR_THROW("Failed to create an instance of ITaskService");
		}

		hr = pService->Connect(_variant_t(), _variant_t(),
			_variant_t(), _variant_t());
		if (FAILED(hr))
		{
			pService->Release();
			pService = NULL;
			CoUninitialize();
			ERROR_THROW("ITaskService::Connect failed");
		}

		hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
		if (FAILED(hr))
		{
			pService->Release();
			pService = NULL;
			CoUninitialize();
			ERROR_THROW("Cannot get Root folder pointer");
		}

		_initialised = true;
	}

	~TaskService() {
		if (pRootFolder != NULL) pRootFolder->Release();
		if (pService != NULL) pService->Release();
		if(_initialised) CoUninitialize();

		pRootFolder = NULL;
		pService = NULL;
		_initialised = false;
	}

	bool IsStarted() {
		return _initialised;
	}

	Task CreateTask(const wstring& taskName) {
		hr = pRootFolder->DeleteTask(_bstr_t(taskName.c_str()), 0);
		if (FAILED(hr)  && !(HRESULT_FACILITY(hr) == FACILITY_WIN32 && HRESULT_CODE(hr) == ERROR_FILE_NOT_FOUND)) {
			ERROR_THROW("Could not delete task");
		}

		ITaskDefinition* pTask = NULL;
		hr = pService->NewTask(0, &pTask);
		IF_ERROR_THROW("Failed to CoCreate an instance of the TaskService class");

		return Task(taskName, pTask, hr);
	}

	void SaveTask(Task& task) {
#ifdef _DEBUG
		BSTR xml;
		hr = task.pTask->get_XmlText(&xml);
		if (FAILED(hr)) {
			cout << "Invalid xml: " << hr << endl;
		} else {
			wcout << xml << endl;
		}

#endif
		IRegisteredTask* pRegisteredTask = NULL;
		hr = pRootFolder->RegisterTaskDefinition(
				_bstr_t(task.taskName.c_str()),
				task.pTask,
				TASK_CREATE_OR_UPDATE,
				_variant_t(L"S-1-5-32-544"),
				_variant_t(),
				TASK_LOGON_GROUP,
				_variant_t(L""),
				&pRegisteredTask);

		IF_ERROR_THROW("Error saving the Task");
		pRegisteredTask->Release();
	}

	// Time format: YYYY-MM-DDTHH:MM:SS
	bool ScheduleEvent(const wstring& path, const wstring& folder, const wstring& time) {
		try {
			Task t = CreateTask(L"SleepSchedulerTask");

			t.SetAuthor(L"SleepScheduler");
			t.AddTimeTrigger(time)->Release();
			t.AddExecutableAction(path, folder)->Release();
			t.SetIdleSettings();
			t.SetLogonType(TASK_LOGON_GROUP, TASK_RUNLEVEL_HIGHEST);
			t.SetStartWhenAvailable(VARIANT_TRUE);
			t.SetStopOnBatteries(VARIANT_FALSE);
			t.SetTimeLimit(L"PT0S");

			SaveTask(t);

			return true;
		} catch(std::exception& ex) {
			cout << ex.what() << endl;
			wcout << HResultToString(hr) << endl;
			return false;
		}
	}
};

void AddDays(chrono::year_month_day& date, unsigned nDays) {
	using namespace std::chrono;
	date = year_month_day(sys_days{ date } + days{ nDays });
}


template <class T, class D>
void AddDays(chrono::time_point<T,D>& time, unsigned nDays) {
	using namespace std::chrono;
	year_month_day ymd{ floor<days>(time) };
	hh_mm_ss hms{ time - floor<days>(time) };
	AddDays(ymd, nDays);
	time = time_point<T, D>{ local_days(ymd) + hms.to_duration() };
}

template <class T, class D>
wstring FormatTime(const chrono::time_point<T,D>& time) {
	using namespace std::chrono;
	year_month_day ymd{ floor<days>(time) };
	hh_mm_ss hms{ floor<minutes>(time) - floor<days>(time) };
	return format(L"{0:%Y}-{0:%m}-{0:%d}T{1:%H}:{1:%M}:{1:%S}", ymd, hms);
}

struct DoubleTime {
	int hour = 0;
	int minute = 0;

	DoubleTime() : hour(0), minute(0) {}
	DoubleTime(int _hour, int _minute) : hour(_hour), minute(_minute) {}

	DoubleTime operator+ (const DoubleTime& t) const {
		return DoubleTime(hour + t.hour + (minute + t.minute) / 60, (minute + t.minute) % 60);
	}
	DoubleTime& operator+= (const DoubleTime& t)  {
		hour += t.hour;
		minute += t.minute;
		hour += minute / 60;
		minute %= 60;
		return *this;
	}

	bool operator== (const DoubleTime& t) const {
		return hour == t.hour && minute == t.minute;
	}
	bool operator> (const DoubleTime& t) const {
		return hour > t.hour || hour == t.hour && minute > t.minute;
	}
	bool operator>= (const DoubleTime& t) const {
		return hour > t.hour || hour == t.hour && minute >= t.minute;
	}
	bool operator< (const DoubleTime& t) const {
		return hour < t.hour || hour == t.hour && minute < t.minute;
	}
	bool operator<= (const DoubleTime& t) const {
		return hour < t.hour || hour == t.hour && minute <= t.minute;
	}

	string to_string() const {
		return format("{:02}:{:02}", hour, minute);
	}

	static const DoubleTime one_day;
	static const DoubleTime one_hour;
	static const DoubleTime one_minute;
	static const DoubleTime zero;
};
const DoubleTime DoubleTime::one_day = DoubleTime(24, 0);
const DoubleTime DoubleTime::one_hour = DoubleTime(1, 0);
const DoubleTime DoubleTime::one_minute = DoubleTime(0, 1);
const DoubleTime DoubleTime::zero = DoubleTime(0, 0);

struct TimeSpan {
	DoubleTime start;
	DoubleTime end;

	bool overlapping(const TimeSpan& t) const {
		if (t < *this) return t.overlapping(*this);
		return end.hour > t.start.hour || end.hour == t.start.hour && end.minute + 1 >= t.start.minute;
	}

	bool operator== (const TimeSpan& t) const {
		return start == t.start;
	}
	bool operator> (const TimeSpan& t) const {
		return start > t.start;
	}
	bool operator< (const TimeSpan& t) const {
		return start < t.start;
	}

	bool contains(const DoubleTime& time) const {
		return time >= start && time <= end;
	}
	
	template<class T>
	bool contains(const chrono::hh_mm_ss<T>& time) const {
		DoubleTime t( time.hours().count(), time.minutes().count() );
		return contains(t);
	}

	string to_string() const {
		return format("{}-{}", start.to_string(), end.to_string());
	}
};

template<class T>
string FormatSpan(const TimeSpan& time, const chrono::hh_mm_ss<T>& now) {
	return format("{:02}:{:02}-{:02}:{:02}{:} ", time.start.hour, time.start.minute, time.end.hour, time.end.minute, time.contains(now) ? " (!!!)" : "");
}

string FormatSpan(const TimeSpan& time) {
	return format("{:02}:{:02}-{:02}:{:02} ", time.start.hour, time.start.minute, time.end.hour, time.end.minute);
}

vector<TimeSpan> spans[7];
int sleepInterval = 0;
const char fileName[] = "schedule.txt";

void ParseFile() {
	ifstream myfile(fileName);
	if (!myfile.is_open()) throw exception("Cannot open schedule file.");

	for (int i = 0; i < 7; i++) {
		spans[i] = vector<TimeSpan>();
	}

	string line;
	getline(myfile, line);
	istringstream ss(line);

	ss >> sleepInterval;

	for (int i = 0; i < 7; i++) {
		getline(myfile, line);
		ss = istringstream(line);

		if (ss.get() != '[') {
			throw exception(format("Schedule file improperly formatted (Line {}) (No opening bracket).", i + 1).c_str());
		}
		if (ss.peek() == ']') continue;

		do {
			TimeSpan ts;
			ss >> ts.start.hour;
			if (ss.get() != ':') throw exception(format("Schedule file improperly formatted (Line {}) (Time formatted incorrectly).", i + 1).c_str());
			ss >> ts.start.minute;
			if (ss.get() != '-') throw exception(format("Schedule file improperly formatted (Line {}) (Time formatted incorrectly).", i + 1).c_str());
			ss >> ts.end.hour;
			if (ss.get() != ':') throw exception(format("Schedule file improperly formatted (Line {}) (Time formatted incorrectly).", i + 1).c_str());
			ss >> ts.end.minute;

			if(ts.start.hour < 0 || ts.start.minute < 0 || ts.end.hour < 0 || ts.end.minute < 0)
				throw exception(format("Cannot have negative time (Line {}) ({}).", i + 1, ts.to_string()).c_str());

			if (ts.start.minute >= 60 || ts.end.minute >= 60)
				throw exception(format("Cannot have minutes over 60 (Line {}) ({}).", i + 1, ts.to_string()).c_str());

			int _i = i;

			while (ts.end < ts.start) ts.end.hour += 24;

			while (ts.start.hour >= 24) {
				_i = (_i + 1) % 7;
				ts.start.hour -= 24;
				ts.end.hour -= 24;
			}

			while (ts.end.hour >= 24) {
				spans[_i].push_back(TimeSpan(ts.start, DoubleTime(23, 59)));
				_i = (_i + 1) % 7;

				ts.start = DoubleTime::zero;
				ts.end.hour -= 24;
			}

			spans[_i].push_back(ts);
		} while (ss.get() == ',');

		ss.unget();

		if(ss.peek() != ']') throw exception(format("Invalid character (Line {}) ({}).", i + 1, (char)ss.get()).c_str());
	}

#ifdef _DEBUG
	cout << "Schedule before merge: " << endl;
	for (int i = 0; i < 7; i++) {
		for (int j = 0; j < spans[i].size(); j++) {
			auto k = spans[i][j];
			cout << FormatSpan(k);
		}
		cout << endl;
	}
	cout << endl;
#endif

	for (int i = 0; i < 7; i++) {
		sort(spans[i].begin(), spans[i].end());

		for (size_t j = 0; j < spans[i].size() - 1;) {
			TimeSpan& a = spans[i][j];
			const TimeSpan& b = spans[i][j + 1];
			if (a.overlapping(b)) {
				a.start = min(a.start, b.start);
				a.end = max(a.end, b.end);
				spans[i].erase(spans[i].begin() + j + 1);
			}
			else j++;
		}
	}
}

void SetPrivilege(const wstring& privilege, bool enable) {
	HANDLE hToken;
	OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken);

	TOKEN_PRIVILEGES tp;
	if (!LookupPrivilegeValue(NULL, privilege.c_str(), &tp.Privileges[0].Luid)) {
		throw exception(format("LookupPrivilegeValue error: {}", GetLastError()).c_str());
	}

	tp.PrivilegeCount = 1;
	tp.Privileges[0].Attributes = enable ? SE_PRIVILEGE_ENABLED : 0;

	if (!AdjustTokenPrivileges(
		hToken,
		FALSE,
		&tp,
		sizeof(TOKEN_PRIVILEGES),
		(PTOKEN_PRIVILEGES)NULL,
		(PDWORD)NULL
	)) {
		throw exception(format("AdjustTokenPrivileges error: {}", GetLastError()).c_str());
	}

	if (GetLastError() == ERROR_NOT_ALL_ASSIGNED) {
		throw exception("The token does not have the specified privilege.");
	}
}

#ifdef _DEBUG
int main()
#else
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR pCmdLine, int nCmdShow)
#endif
{
	using namespace std::chrono;

	wchar_t fileName[256];
	wchar_t execPath[256];

	if (GetModuleFileName(NULL, fileName, sizeof(fileName) / sizeof(wchar_t)) == 0) {
		cout << "Cannot get application path." << endl;
		return 1;
	}

	memcpy(execPath, fileName, sizeof(fileName));
	PathCchRemoveFileSpec(execPath, sizeof(fileName) / sizeof(wchar_t));

	try {
		SetPrivilege(SE_SHUTDOWN_NAME, true);
	}
	catch (const std::exception& e) {
		cout << "Cannot gain shutdown permission:" << endl;
		cout << e.what() << endl;
		return 1;
	}

	ParseFile();

	int tswd = -1;
	TimeSpan* ts = NULL;
	chrono::local_time<chrono::system_clock::duration> tp = zoned_time{ current_zone(), system_clock::now() }.get_local_time();
	int wd;
	hh_mm_ss<minutes> t{ floor<minutes>(tp) - floor<days>(tp) };

#ifdef _DEBUG
	cout << "Schedule after merge: " << endl;
	for (int i = 0; i < 7; i++) {
		for (int j = 0; j < spans[i].size(); j++) {
			cout << FormatSpan(spans[i][j], t);
		}
		cout << endl;
	}
	cout << endl;
#endif

	do {
		tp = zoned_time{ current_zone(), system_clock::now() }.get_local_time();
		// Sun - Sat, 0-6
		wd = weekday{ floor<days>(tp) }.c_encoding();
		t = hh_mm_ss{ floor<minutes>(tp) - floor<days>(tp) };
		
		if (ts == NULL || wd != tswd || !ts->contains(t)) {
			ts = NULL;
			for (size_t i = 0; i < spans[wd].size(); i++) {
				if (spans[wd][i].contains(t)) {
					ts = &(spans[wd][i]);
					tswd = wd;
					break;
				}
			}
		}
		else {
#ifdef _DEBUG
			cout << "Sleep" << endl;
			getchar();
#else
			SetSuspendState(false, false, false);
			Sleep(sleepInterval);
#endif
		}
	} while (ts != NULL);

	bool set = false;
	year_month_day next_date = year_month_day(floor<days>(tp));
	DoubleTime next_start;

	DoubleTime t2{ t.hours().count(), t.minutes().count() };
	for (size_t j = 0; j < spans[wd].size(); j++) {
		if (spans[wd][j].start < t2) continue;
		next_start = spans[wd][j].start;
		set = true;
		break;
	}

	if (!set) {
		for (int i = 1; i < 7; i++) {
			if (spans[(wd + i) % 7].empty()) continue;
			AddDays(next_date, i);
			next_start = spans[wd][0].start;
			set = true;
			break;
		}
	}

	if (!set) return 0;

	time_point next_time = local_days(next_date) + hours(next_start.hour) + minutes(next_start.minute);

	TaskService tserv;
	bool result = tserv.ScheduleEvent(L'"' + wstring(fileName) + L'"', wstring(execPath), FormatTime(next_time));

#ifdef _DEBUG
	wcout << L"Set next time: " << FormatTime(next_time) << endl;
	getchar();
#endif
	return result ? 0 : 1;
}