#ifndef __QUEUE_H__
#define __QUEUE_H__

#include <list>
#include <thread>
#include <mutex>
#include <afxstr.h>
#include <condition_variable>

struct ProgressCmd
{
	int type;
	int id;
	int arg1;
	int arg2;
};

template<class T>
class MQ
{
public:
	MQ() {}

	~MQ() {}

public:
	static MQ<T>& getInstance()
	{
		static MQ<T> self;
		return self;
	}

public:
	void Push(T data)
	{
		std::lock_guard<std::mutex> lock(m_lock);
		m_queue.push_back(data);
		m_cv.notify_all();
	}

	T Pop()
	{
		std::unique_lock<std::mutex> lock(m_lock);
		while (m_queue.empty())
		{
			m_cv.wait(lock);
		}
		T data = m_queue.front();
		m_queue.pop_front();
		return data;
	}

private:
	std::mutex m_lock;
	std::condition_variable m_cv;
	std::list<T> m_queue;
};


class Queue
{
public:
	Queue() {}

	~Queue() {}

public:
	static Queue& getInstance()
	{
		static Queue self;
		return self;
	}

public:
	void Push(const CString& data)
	{
		std::lock_guard<std::mutex> lock(m_lock);
		m_queue.push_back(data);
		m_cv.notify_all();
	}

	CString Pop() 
	{ 
		std::unique_lock<std::mutex> lock(m_lock);
		while (m_queue.empty())
		{
			m_cv.wait(lock);
		}
		CString data = m_queue.front();
		m_queue.pop_front();
		return data;
	}

private:
	std::mutex m_lock;
	std::condition_variable m_cv;
	std::list<CString> m_queue;
};

#endif