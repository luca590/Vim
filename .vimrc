set ruler laststatus=2 number title hlsearch 
set hlsearch 
set incsearch

set showmatch
set enc=utf-8
set fenc=utf-8
set termencoding=utf-8 
"set smartindent
"set autoindent
syntax on
colorscheme desert

autocmd InsertEnter * :set number
autocmd InsertLeave * :set relativenumber
set tabstop=4
imap jj <esc>
nnoremap S <C-W><C-W>

" ---------- Vim-plug, vim plugin manager auto install -------
"if empty(glob('~/.vim/autoload/plug.vim'))
"  silent !curl -fLo ~/.vim/autoload/plug.vim --create-dirs
"    \ https://raw.githubusercontent.com/junegunn/vim-plug/master/plug.vim
"  autocmd VimEnter * PlugInstall --sync | source $MYVIMRC
"endif

"call plug#begin('~/.vim/plugged')

"Plug 'elixir-editors/vim-elixir'	"elixir vim plugin
"Plug 'tomlion/vim-solidity'		"solidity vim plugin
"Plug 'vim-scripts/c.vim'		"C plugin

call plug#end()

