#include "stdafx.h"
#include "get_rest.h"
#include <iostream>
#include <sstream>

#if  defined(__WIN32__) || defined(_WIN32)

static int _haveInitializedWinsock = 0;
int initializeWinsockIfNecessary(void)
{
	/* We need to call an initialization routine before
	 * we can do anything with winsock.  (How fucking lame!):
	 */
	
	WSADATA	wsadata;

	if (!_haveInitializedWinsock) {
		if (WSAStartup(MAKEWORD(2,2),  &wsadata) != 0) {
			return 0; /* error in initialization */
		}
	    	if (wsadata.wVersion != MAKEWORD(2,2)) {
	        	WSACleanup();
				return 0; /* desired Winsock version was not available */
		}
		_haveInitializedWinsock = 1;
	}

	return 1;
}

void uninitializeWinsockIfNecessary(void)
{
	if(_haveInitializedWinsock){
		WSACleanup();
		_haveInitializedWinsock = 0;
	}
}
#else
/* Unix */
#endif //#if  defined(__WIN32__) || defined(_WIN32)


bool sky_rest::parse_url(const std::string url, std::string& host, short& port, std::string& path)
{
	if(url.empty()) return false;

	std::string tmp, host_port;
	if(url.compare(0, 7, "http://") == 0)
		tmp = url.substr(7);
	else
		tmp = url;
	
	int pos =  tmp.find_first_of('/');
	if(pos > 0){
		host_port = tmp.substr(0, pos);
		path = tmp.substr(pos);
	}
	else{
		host_port = tmp;
		path = "/";
	}
	
	int pos1 = host_port.find_first_of(':');
	if(pos1 > 0){
		host = host_port.substr(0, pos1);
		std::string p = host_port.substr(pos1 + 1, pos - pos1 -1);
		port = atoi(p.c_str());
	}
	else{
		host = host_port;
		port = 80;
	}
	
	return true;
}

int sky_rest::connect_to_server(const std::string host, const short port, long connect_time_out)
{
	struct hostent* phe;
	struct sockaddr_in sin;
	int s;
	struct   timeval time_out;

	memset(&sin, 0, sizeof(sin));
	sin.sin_family = AF_INET;

	if((phe = gethostbyname(host.c_str())))
		memcpy(&sin.sin_addr, phe->h_addr, phe->h_length);
	else
		sin.sin_addr.s_addr = inet_addr(host.c_str());
	//std::cout << inet_ntoa(sin.sin_addr) << std::endl;
	sin.sin_port = htons(port);
	s = socket(PF_INET, SOCK_STREAM, 0);
	if(s < 0) return -1;
	int max_s;
#if  defined(__WIN32__) || defined(_WIN32)
	unsigned long nonbloking_mode =1;// 0: bloking mode; !0: nonbloking mode
	if(ioctlsocket(s, FIONBIO, &nonbloking_mode)){
		CloseSock(s);
		return -1;
	}
	max_s = 0;
#else
	int flags = fcntl(s, F_GETFL, 0);
	if(fcntl(s, F_SETFL, flags | O_NONBLOCK)){
		CloseSock(s);
		return -1;
	}
	max_s = s+1;
#endif
	if(connect(s , (struct sockaddr*) &sin, sizeof(sin)) == -1 && REST_EWOULDBLOCK != rest_getLastError()){
		CloseSock(s);
		return -1;
	}
	time_out.tv_sec = connect_time_out;
	time_out.tv_usec = 0;
	fd_set r;
	int ret;
	
	while(1){
		FD_ZERO(&r);
		FD_SET(s, &r);
		ret = select(max_s, 0, &r, 0, &time_out);
		if( ret <= 0 ) 
		{
			CloseSock(s);
			return -1;
		}
		if(FD_ISSET(s, &r)) break;
	}
#if  defined(__WIN32__) || defined(_WIN32)
	nonbloking_mode = 0; 
	ret   =   ioctlsocket(s, FIONBIO, &nonbloking_mode); 
	if(ret==SOCKET_ERROR){ 
		CloseSock(s); 
		return -1;
	} 
#else
	flags = fcntl(s, F_GETFL, 0);
	if(fcntl(s, F_SETFL, flags & ~ O_NONBLOCK)){
		CloseSock(s);
		return -1;
	}
#endif

	return s;
}

int sky_rest::send_request(int s, const std::string path, const std::string token, const std::string host, const short port)
{
	std::stringstream request;
	if(token.empty())
		request << "GET " << path << " HTTP/1.1\r\n"
				<< "Host: " << host << ":" << port << "\r\n"
				<< "Connection: keep-alive\r\n"
				<< "\r\n";
	else
		request << "GET " << path << " HTTP/1.1\r\n"
				<< "Host: " << host << ":" << port << "\r\n"
				<< "Cookie: csst_token=" << token << "\r\n"
				<< "Connection: keep-alive\r\n"
				<< "\r\n";
	std::string s_req = request.str();
	int ret = send(s, s_req.c_str(), s_req.length(), 0);
	//shutdown(s, REST_SEND);
	return ret;
}

int sky_rest::send_post_request(int s, const std::string path, const std::string data, const std::string token, const std::string host, const short port)
{
	std::stringstream request;
	if(token.empty())
		request << "POST " << path << " HTTP/1.1\r\n"
				<< "Host: " << host << ":" << port << "\r\n"
				<< "Content-Type: application/x-www-form-urlencoded\r\n"
				<< "Content-Length: " << data.length() << "\r\n"
				<< "Connection: keep-alive\r\n"
				<< "\r\n"
				<< data;
	else
		request << "POST " << path << " HTTP/1.1\r\n"
				<< "Host: " << host << ":" << port << "\r\n"
				<< "Cookie: csst_token=" << token << "\r\n"
				<< "Content-Type: application/x-www-form-urlencoded\r\n"
				<< "Content-Length: " << data.length() << "\r\n"
				<< "Connection: keep-alive\r\n"
				<< "\r\n"
				<< data;
	int ret = send(s, request.str().c_str(), request.str().length(), 0);
	//shutdown(s, REST_SEND);
	return ret;
}

int sky_rest::get_response(int s, std::string& result)
{
	std::stringstream ss;
	char buff[100*1024];
	int received = 0, ret;
	memset(buff, 0, 100*1024);
	while((ret = recv(s, buff+received, 100*1024-received, 0)) > 0){
		ss << buff;
		received += ret;
		//memset(buff, 0, 10*1024);
		if (!check_receiv_complete(buff)) {
			break;
		}
	}
	if(received > 0) result = buff;//ss.str();
	return received;
}

int sky_rest::check_receiv_complete(const char* buf)
{
	int ret = -1;
	int contentLen = -1;
	std::istringstream input;
	input.str(buf);
	for (std::string line; std::getline(input, line); ) {
		std::string::size_type n = line.find("Content-Length");
		if (n == std::string::npos) {
			continue;
		} else {
			std::string sLen = line.substr(15);
			contentLen = atoi(sLen.c_str());
			std::string body;
			get_response_body(buf, body);
			if (body.length() == contentLen) {
				return 0;
			}
		}
	}
	return -1;
}

int sky_rest::get_response_body(const std::string response, std::string& body)
{
	int state = -1;
	int pos_header, pos_state;
	if(((pos_state = response.find("\r\n", 0, 2)) > 0)){
		std::string ss = response.substr(8, pos_state - 8);
		const char* s = ss.c_str();
		while(*s == ' ' || *s == '\t') s++;
		std::string n = s;
		pos_state = n.find(' ');
		if(pos_state < 0) pos_state = n.find('\t');
		if(pos_state > 0){
			std::string s_state = n.substr(0, pos_state);
			state = atoi(s_state.c_str());
		}
	}
	if((pos_header = response.find("\r\n\r\n", 0, 4)) > 0){
		body = response.substr(pos_header + 4);
	}
	return state;
}

int sky_rest::get_rest(const std::string url, std::string& body, long connect_time_out, const std::string token)
{
	int result = -1;
	int state = result;
	std::string host, path, response;
	short port;
	int s;
	body = "";
	if(parse_url(url, host, port, path)){
		s = connect_to_server(host, port, connect_time_out);
		if(s > 0){
			if(send_request(s, path, token, host, port) > 0){
				if(get_response(s, response) > 0){
					state = get_response_body(response, body);
					//if(state >= 200 && state < 300)
						result = state;
				}
			}
			CloseSock(s);
		}
	}
	return result;
}

int sky_rest::post_rest(const std::string url, const std::string data, std::string& body, long connect_time_out, const std::string token)
{
	int result = -1;
	int state = result;
	std::string host, path, response;
	short port;
	int s;
	body = "";
	if(parse_url(url, host, port, path)){
		s = connect_to_server(host, port, connect_time_out);
		if(s > 0){
			if(send_post_request(s, path, data, token, host, port) > 0){
				if(get_response(s, response) > 0){
					state = get_response_body(response, body);
					if(state >= 200 && state < 300)
						result = 0;
				}
			}
			CloseSock(s);
		}
	}
	return result;
}
