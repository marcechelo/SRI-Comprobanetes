

bind = "127.0.0.1:9000"
max_requests = 1000
worker_class = 'gevent'

workers = 2     # the number of recommended workers is 2 * number of CPUs +1
