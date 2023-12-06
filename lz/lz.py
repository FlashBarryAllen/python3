def compress_lz77(data, window_size, lookahead_buffer_size):
    dictionary = ""
    compressed_data = []
    i = 0
 
    while i < len(data):
        search_window_start = max(0, i - window_size)
        search_window_end = i
        search_window = data[search_window_start:search_window_end]
        match_length = 0
        match_index = 0
 
        for j in range(lookahead_buffer_size):
            if i + j >= len(data):
                break
            current_string = data[i:i+j+1]
 
            if current_string in search_window and len(current_string) > match_length:
                match_length = len(current_string)
                match_index = search_window.index(current_string)
 
        if match_length > 0:
            compressed_data.append((match_index, match_length))
            i += match_length
        else:
            compressed_data.append((0, 0))
            i += 1
 
        dictionary += data[i-1:i+match_length]
 
    return compressed_data
 
def decompress_lz77(compressed_data, window_size):
    dictionary = "a"
    decompressed_data = ""
 
    for (match_index, match_length) in compressed_data:
        if match_length == 0:
            print(dictionary[-1])
            decompressed_data += dictionary[-1]
            dictionary += dictionary[-1]
        else:
            start_index = len(dictionary) - window_size
            dictionary += dictionary[start_index+match_index:start_index+match_index+match_length]
            decompressed_data += dictionary[-match_length:]
 
    return decompressed_data
 
# 示例用法
data = "ababcbababaaaaaa"
window_size = 5
lookahead_buffer_size = 3
 
compressed_data = compress_lz77(data, window_size, lookahead_buffer_size)
decompressed_data = decompress_lz77(compressed_data, window_size)
 
print("原始数据:", data)
print("压缩后的数据:", compressed_data)
print("解压缩后的数据:", decompressed_data)